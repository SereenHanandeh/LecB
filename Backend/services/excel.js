const { getPlanContext, saveAssignments, getDutyPool } = require("./plan.js");

const seededShuffle = require("../utils/rng.js");

const { getPeriodRank, normalizePeriod } = require("../utils/distribution.js");

const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

// ============================================================
// Config

const ALLOW_SAME_PERIOD_MULTIPLE_PROFESSORS = true;

// ============================================================
// Excel Import
// ============================================================

function importExcel(filePath) {
  const workbook = xlsx.readFile(filePath, {
    cellDates: true,
  });

  console.log("========================================");
  console.log("📚 EXCEL SHEETS DEBUG");
  console.log("========================================");
  console.log("📚 All sheets:", workbook.SheetNames);

  for (const name of workbook.SheetNames) {
    const testSheet = workbook.Sheets[name];

    const testRows = xlsx.utils.sheet_to_json(testSheet, {
      range: 2,
      defval: null,
      raw: false,
    });

    console.log(`📄 Sheet "${name}" => ${testRows.length} rows`);
  }

  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error("Excel file does not contain any sheets.");
  }

  console.log("📄 Reading worksheet:", sheetName);

  const sheet = workbook.Sheets[sheetName];

  const rawRows = xlsx.utils.sheet_to_json(sheet, {
    range: 2,
    defval: null,
    raw: false,
  });

  // ============================================================
  // Excel Statistics
  // ============================================================

  console.log("📊 EXCEL IMPORT DEBUG");
  console.log("========================================");
  console.log("📊 Excel rows read:", rawRows.length);

  const excelDates = {};

  for (const row of rawRows) {
    const dateValue =
      row["التاريخ ميلادي"] ??
      row["Date"] ??
      row["DATE"] ??
      row["date"] ??
      null;

    const normalizedDate = dateISO(dateValue);

    if (!excelDates[normalizedDate]) {
      excelDates[normalizedDate] = 0;
    }

    excelDates[normalizedDate]++;
  }

  console.log("📋 Excel dates:");
  console.log(excelDates);

  if (!rawRows.length) {
    return [];
  }

  // ============================================================
  // Normalize Rows
  // ============================================================

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

    const crn = row["CRN"] ?? row["crn"] ?? "";

    const courseName =
      row["اسم المقرر"] ?? row["Course Name"] ?? row["course_name"] ?? "";

    const lecture =
      row["المحاضرة المباشرة"] ?? row["Lecture"] ?? row["lecture"] ?? "";

    const timeFrom = row["من"] ?? row["From"] ?? row["from"] ?? "";

    const timeTo = row["إلى"] ?? row["To"] ?? row["to"] ?? "";

    return {
      crn: String(crn).trim(),

      professor_name: String(professor).trim(),

      date,

      period_label: String(period).trim(),

      course_name: String(courseName).trim(),

      lecture: String(lecture).trim(),

      time_from: String(timeFrom).trim(),

      time_to: String(timeTo).trim(),

      required_supervisors: 1,

      original: row,

      excel_row: index + 3,
    };
  });

  // ============================================================
  // Validate Rows
  // ============================================================

  const validRows = rows.filter((row) => {
    return row.crn && row.professor_name && row.date && row.period_label;
  });

  console.log(`✅ Valid Excel rows: ${validRows.length}/${rows.length}`);

  if (validRows.length) {
    console.log("📦 First normalized Excel row:", validRows[0]);
  }

  return validRows;
}

// ============================================================
// Date Helpers
// ============================================================

function dateISO(value) {
  if (!value) return "";

  if (typeof value === "string") {
    const str = value.trim();

    if (!str) return "";

    const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);

    if (isoMatch) {
      return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
    }

    const dateMatch = str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);

    if (dateMatch) {
      const day = String(dateMatch[1]).padStart(2, "0");

      const month = String(dateMatch[2]).padStart(2, "0");

      const year = dateMatch[3];

      return `${year}-${month}-${day}`;
    }

    const d = new Date(str);

    if (!Number.isNaN(d.getTime())) {
      const year = d.getFullYear();

      const month = String(d.getMonth() + 1).padStart(2, "0");

      const day = String(d.getDate()).padStart(2, "0");

      return `${year}-${month}-${day}`;
    }
  }

  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      return "";
    }

    const year = value.getFullYear();

    const month = String(value.getMonth() + 1).padStart(2, "0");

    const day = String(value.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  const d = new Date(value);

  if (!Number.isNaN(d.getTime())) {
    const year = d.getFullYear();

    const month = String(d.getMonth() + 1).padStart(2, "0");

    const day = String(d.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  return String(value).trim();
}

// ============================================================
// Slot Helpers
// ============================================================

function getSupervisorSlotKey(day, period) {
  return `${day}|${normalizePeriod(period)}`;
}

function isValidSequentialPeriodAttachment(supervisor, day, newRanks) {
  const cleanRanks = (newRanks || []).filter(
    (rank) => rank !== null && rank !== undefined,
  );

  if (!cleanRanks.length) {
    // لا يوجد Rank معروف لهذه الفترة، القاعدة لا تنطبق
    return true;
  }

  const existingRanks = supervisor.byDayPeriodRanks?.[day];

  let uniqueNew = [...new Set(cleanRanks)];

  // ✅ الفترات الموجودة عند المشرف مسبقًا لا تُعد فترات جديدة
  if (ALLOW_SAME_PERIOD_MULTIPLE_PROFESSORS && existingRanks?.size) {
    uniqueNew = uniqueNew.filter((rank) => !existingRanks.has(rank));
  }

  // كل الفترات مكررة (نفس فترات المشرف الحالية) => مسموح
  if (!uniqueNew.length) {
    return true;
  }

  const sortedNew = uniqueNew.sort((a, b) => a - b);

  // الفترات الجديدة نفسها يجب أن تكون متسلسلة بدون فراغ داخلي
  for (let i = 1; i < sortedNew.length; i++) {
    if (sortedNew[i] !== sortedNew[i - 1] + 1) {
      return false;
    }
  }

  if (!existingRanks || existingRanks.size === 0) {
    // أول تعيين لهذا المشرف في هذا اليوم
    return true;
  }

  const existingMax = Math.max(...existingRanks);

  // الفترات الجديدة يجب أن تبدأ مباشرة بعد آخر فترة موجودة
  return sortedNew[0] === existingMax + 1;
}

function addDaysISO(day, amount) {
  const match = String(day).match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) return "";

  const year = Number(match[1]);
  const month = Number(match[2]);
  const date = Number(match[3]);

  const d = new Date(Date.UTC(year, month - 1, date));

  d.setUTCDate(d.getUTCDate() + amount);

  return [
    d.getUTCFullYear(),
    String(d.getUTCMonth() + 1).padStart(2, "0"),
    String(d.getUTCDate()).padStart(2, "0"),
  ].join("-");
}

// ============================================================
// Consecutive Day Rule
// ============================================================

function hasConsecutiveDayConflict(supervisor, professorBundles) {
  const existingDays = new Set(
    Object.keys(supervisor.byDay || {}).filter(
      (day) => Number(supervisor.byDay[day]) > 0,
    ),
  );

  const newDays = [
    ...new Set(
      professorBundles.map((bundle) => dateISO(bundle.date)).filter(Boolean),
    ),
  ];

  for (const day of newDays) {
    const previousDay = addDaysISO(day, -1);

    const nextDay = addDaysISO(day, 1);

    if (existingDays.has(previousDay) || existingDays.has(nextDay)) {
      return true;
    }
  }

  return false;
}

// ============================================================
// Period Sorting
// ============================================================

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

function attachGroupToSupervisor(supervisor, group, result) {
  if (!supervisor || !group) {
    return false;
  }

  const id = Number(supervisor.id);

  if (!Number.isFinite(id)) {
    return false;
  }

  const groupId = Number(group.id);

  if (!Number.isFinite(groupId)) {
    return false;
  }

  if (supervisor.assignedGroups.has(groupId)) {
    return false;
  }

  supervisor.assignedGroups.add(groupId);

  result.push({
    session_group_id: groupId,

    crn: group.crn,

    // ✅ اسم المقرر
    course_name: group.course_name ?? "",

    // ✅ الأستاذ
    professor: group.professor_name || group.professor || "",
    professor_id: group.professor_id ?? null,

    // ✅ رقم القاعة
    room_number: group.room_number ?? null,

    // ✅ التاريخ
    date: dateISO(group.date),

    // ✅ الفترة
    period: normalizePeriod(group.period_label),

    // ✅ وقت البداية والنهاية
    time_from: group.time_from ?? "",
    time_to: group.time_to ?? "",

    // ✅ المشرف
    supervisor_id: id,
  });

  return true;
}

// ============================================================
// Assign Complete Bundle
// ============================================================

function assignBundleToSupervisor(
  supervisor,
  bundle,
  result,
  bundleAssignments,
) {
  if (!supervisor || !bundle) {
    return false;
  }

  const groups = bundle.groups || [];

  if (!groups.length) {
    return false;
  }

  const representative = groups[0];

  const day = dateISO(representative.date);

  const period = normalizePeriod(representative.period_label);

  const periodRank = getPeriodRank(period);

  if (!day || !period) {
    return false;
  }

  const slotKey = getSupervisorSlotKey(day, period);

  if (bundleAssignments.has(bundle.key)) {
    const assignedSupervisorId = bundleAssignments.get(bundle.key);

    if (Number(assignedSupervisorId) !== Number(supervisor.id)) {
      return false;
    }

    return true;
  }

  if (!supervisor.byDayPeriods[day]) {
    supervisor.byDayPeriods[day] = new Set();
  }

  if (!supervisor.byDayPeriodRanks[day]) {
    supervisor.byDayPeriodRanks[day] = new Set();
  }

  if (!ALLOW_SAME_PERIOD_MULTIPLE_PROFESSORS) {
    if (supervisor.occupiedSlots && supervisor.occupiedSlots.has(slotKey)) {
      return false;
    }

    if (supervisor.byDayPeriods[day].has(period)) {
      return false;
    }

    if (
      periodRank !== null &&
      periodRank !== undefined &&
      supervisor.byDayPeriodRanks[day].has(periodRank)
    ) {
      return false;
    }
  }

  for (const group of groups) {
    if (!group || group.id === null || group.id === undefined) {
      return false;
    }

    const groupId = Number(group.id);

    if (supervisor.assignedGroups.has(groupId)) {
      return false;
    }
  }

  for (const group of groups) {
    const added = attachGroupToSupervisor(supervisor, group, result);

    if (!added) {
      return false;
    }
  }

  if (!supervisor.byDay[day]) {
    supervisor.byDay[day] = 0;
  }

  supervisor.byDay[day] += 1;

  supervisor.total += 1;

  supervisor.lastDay = day;

  supervisor.byDayPeriods[day].add(period);

  if (periodRank !== null && periodRank !== undefined) {
    supervisor.byDayPeriodRanks[day].add(periodRank);
  }

  supervisor.occupiedSlots.add(slotKey);

  bundleAssignments.set(bundle.key, supervisor.id);

  return true;
}

// ============================================================
// Check Professor
// ============================================================

function canSupervisorTakeProfessor(
  supervisor,
  professorBundles,
  options = {},
) {
  const { relaxed = false } = options;

  if (!supervisor || !professorBundles?.length) {
    return false;
  }

  const professorSlots = new Set();

  for (const bundle of professorBundles) {
    const representative = bundle.groups?.[0];

    if (!representative) {
      return false;
    }

    const day = dateISO(representative.date);

    const period = normalizePeriod(representative.period_label);

    if (!day || !period) {
      return false;
    }

    const slotKey = getSupervisorSlotKey(day, period);

    if (professorSlots.has(slotKey)) {
      return false;
    }

    professorSlots.add(slotKey);
  }

  for (const bundle of professorBundles) {
    const representative = bundle.groups?.[0];

    const day = dateISO(representative.date);

    const period = normalizePeriod(representative.period_label);

    const slotKey = getSupervisorSlotKey(day, period);

    if (!ALLOW_SAME_PERIOD_MULTIPLE_PROFESSORS) {
      if (supervisor.occupiedSlots && supervisor.occupiedSlots.has(slotKey)) {
        return false;
      }

      if (supervisor.byDayPeriods?.[day]?.has(period)) {
        return false;
      }

      const periodRank = getPeriodRank(period);

      if (
        periodRank !== null &&
        periodRank !== undefined &&
        supervisor.byDayPeriodRanks?.[day]?.has(periodRank)
      ) {
        return false;
      }
    }
  }

  // --------------------------------------------------------
  // Sequential Period Attachment Rule (Hard Rule)
  // تُتجاهل فقط في وضع Relaxed (ملاذ أخير لضمان تغطية كل الأساتذة)
  // --------------------------------------------------------

  if (!relaxed) {
    const newRanksByDay = new Map();

    for (const bundle of professorBundles) {
      const representative = bundle.groups?.[0];

      const day = dateISO(representative.date);
      const period = normalizePeriod(representative.period_label);
      const rank = getPeriodRank(period);

      if (!day || rank === null || rank === undefined) {
        continue;
      }

      if (!newRanksByDay.has(day)) {
        newRanksByDay.set(day, []);
      }

      newRanksByDay.get(day).push(rank);
    }

    for (const [day, ranks] of newRanksByDay) {
      if (!isValidSequentialPeriodAttachment(supervisor, day, ranks)) {
        return false;
      }
    }

    if (hasConsecutiveDayConflict(supervisor, professorBundles)) {
      return false;
    }
  }

  return true;
}

// ============================================================
// Consecutive Period Score
// ============================================================

function getProfessorConsecutiveScore(supervisor, professorBundles) {
  let score = 0;

  const existingByDay = new Map();

  for (const [day, ranks] of Object.entries(
    supervisor.byDayPeriodRanks || {},
  )) {
    existingByDay.set(day, new Set(ranks));
  }

  const professorByDay = new Map();

  for (const bundle of professorBundles) {
    const representative = bundle.groups?.[0];

    if (!representative) {
      continue;
    }

    const day = dateISO(representative.date);

    const period = normalizePeriod(representative.period_label);

    const rank = getPeriodRank(period);

    if (rank === null || rank === undefined) {
      continue;
    }

    if (!professorByDay.has(day)) {
      professorByDay.set(day, new Set());
    }

    professorByDay.get(day).add(rank);
  }

  for (const [day, candidateRanks] of professorByDay) {
    const existingRanks = existingByDay.get(day) || new Set();

    for (const rank of candidateRanks) {
      if (existingRanks.has(rank - 1)) {
        score += 100;
      }

      if (existingRanks.has(rank + 1)) {
        score += 60;
      }

      if (candidateRanks.has(rank - 1)) {
        score += 50;
      }

      if (candidateRanks.has(rank + 1)) {
        score += 30;
      }

      if (existingRanks.size > 1) {
        const sorted = [...existingRanks].sort((a, b) => a - b);

        const minRank = sorted[0];

        const maxRank = sorted[sorted.length - 1];

        if (rank > minRank && rank < maxRank && !existingRanks.has(rank)) {
          score += 80;
        }
      }

      if (existingRanks.has(rank - 1) && existingRanks.has(rank - 2)) {
        score += 40;
      }
    }
  }

  return score;
}

// ============================================================
// Period Continuity Score
// ============================================================

function getPeriodContinuityScore(supervisor, professorBundles) {
  let score = 0;

  const candidateByDay = new Map();

  for (const bundle of professorBundles) {
    const representative = bundle.groups?.[0];

    if (!representative) {
      continue;
    }

    const day = dateISO(representative.date);

    const period = normalizePeriod(representative.period_label);

    const rank = getPeriodRank(period);

    if (!day || rank === null || rank === undefined) {
      continue;
    }

    if (!candidateByDay.has(day)) {
      candidateByDay.set(day, new Set());
    }

    candidateByDay.get(day).add(rank);
  }

  function longestContinuousRun(set) {
    const ranks = [...set].sort((a, b) => a - b);

    if (!ranks.length) {
      return 0;
    }

    let best = 1;
    let current = 1;

    for (let i = 1; i < ranks.length; i++) {
      if (ranks[i] === ranks[i - 1] + 1) {
        current++;

        if (current > best) {
          best = current;
        }
      } else {
        current = 1;
      }
    }

    return best;
  }

  function countInternalGaps(set) {
    const ranks = [...set].sort((a, b) => a - b);

    if (ranks.length < 2) {
      return 0;
    }

    const minRank = ranks[0];

    const maxRank = ranks[ranks.length - 1];

    const expectedCount = maxRank - minRank + 1;

    const actualCount = new Set(ranks).size;

    return Math.max(0, expectedCount - actualCount);
  }

  function hasAdjacentPeriod(existingRanks, rank) {
    return existingRanks.has(rank - 1) || existingRanks.has(rank + 1);
  }

  for (const [day, candidateRanks] of candidateByDay) {
    const existingRanks = new Set(supervisor.byDayPeriodRanks?.[day] || []);

    const beforeRun = longestContinuousRun(existingRanks);

    const beforeGaps = countInternalGaps(existingRanks);

    const afterRanks = new Set(existingRanks);

    for (const rank of candidateRanks) {
      afterRanks.add(rank);
    }

    const afterRun = longestContinuousRun(afterRanks);

    const afterGaps = countInternalGaps(afterRanks);

    for (const rank of candidateRanks) {
      const fillsGap =
        existingRanks.has(rank - 1) && existingRanks.has(rank + 1);

      if (fillsGap) {
        score += 1000000;
      }
    }

    const gapReduction = beforeGaps - afterGaps;

    if (gapReduction > 0) {
      score += gapReduction * 200000;
    }

    const runGrowth = afterRun - beforeRun;

    if (runGrowth > 0) {
      score += runGrowth * 100000;
    }

    for (const rank of candidateRanks) {
      if (existingRanks.has(rank - 1)) {
        score += 30000;
      }
    }

    for (const rank of candidateRanks) {
      if (existingRanks.has(rank + 1)) {
        score += 25000;
      }
    }

    for (const rank of candidateRanks) {
      const connectsBothSides =
        existingRanks.has(rank - 1) && existingRanks.has(rank + 1);

      if (connectsBothSides) {
        score += 500000;
      }
    }

    for (const rank of candidateRanks) {
      if (existingRanks.has(rank - 1) && existingRanks.has(rank - 2)) {
        score += 20000;
      }

      if (existingRanks.has(rank + 1) && existingRanks.has(rank + 2)) {
        score += 18000;
      }
    }

    const candidateRun = longestContinuousRun(candidateRanks);

    if (candidateRun > 1) {
      score += (candidateRun - 1) * 10000;
    }

    for (const rank of candidateRanks) {
      if (hasAdjacentPeriod(existingRanks, rank)) {
        score += 5000;
      }
    }

    const newGaps = afterGaps - beforeGaps;

    if (newGaps > 0) {
      score -= newGaps * 100000;
    }

    if (afterRanks.size > 0 && afterGaps === 0) {
      score += 50000;
    }
  }

  return score;
}

// ============================================================
// Continuity Priority
// ============================================================

function getContinuityPriority(supervisor, professorBundles) {
  const continuityScore = getPeriodContinuityScore(
    supervisor,
    professorBundles,
  );

  const consecutiveScore = getProfessorConsecutiveScore(
    supervisor,
    professorBundles,
  );

  return continuityScore + consecutiveScore * 10;
}

// ============================================================
// Daily Load
// ============================================================

function getProfessorDailyLoad(supervisor, professorBundles) {
  let dailyLoad = 0;

  for (const bundle of professorBundles) {
    const day = dateISO(bundle.date);

    dailyLoad += Number(supervisor.byDay?.[day] || 0);
  }

  return dailyLoad;
}

// ============================================================
// Minimum Period Helpers
// ============================================================

function hasReachedAllMinimums(
  cand,
  selectedSupervisorIds,
  minimumPeriodsEnabled,
  minimumPeriods,
) {
  if (!minimumPeriodsEnabled) {
    return true;
  }

  for (const supervisorId of selectedSupervisorIds) {
    const supervisor = cand[Number(supervisorId)];

    if (!supervisor) {
      continue;
    }

    if (Number(supervisor.total || 0) < minimumPeriods) {
      return false;
    }
  }

  return true;
}

function getMinimumInfo(
  supervisor,
  workload,
  minimumPeriodsEnabled,
  minimumPeriods,
) {
  const currentTotal = Number(supervisor.total || 0);

  const projectedTotal = currentTotal + workload;

  if (!minimumPeriodsEnabled) {
    return {
      needsMinimum: false,
      deficit: 0,
      projectedDeficit: 0,
      reachesMinimum: false,
    };
  }

  const deficit = Math.max(0, minimumPeriods - currentTotal);

  const projectedDeficit = Math.max(0, minimumPeriods - projectedTotal);

  const needsMinimum = currentTotal < minimumPeriods;

  const reachesMinimum =
    currentTotal < minimumPeriods && projectedTotal >= minimumPeriods;

  return {
    needsMinimum,
    deficit,
    projectedDeficit,
    reachesMinimum,
  };
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
  options = {},
) {
  const id = Number(supervisorId);

  if (!Number.isFinite(id)) {
    return false;
  }

  if (!cand[id]) {
    return false;
  }

  const supervisor = cand[id];

  if (!canSupervisorTakeProfessor(supervisor, professorBundles, options)) {
    return false;
  }

  const previousResultLength = result.length;

  const previousTotal = supervisor.total;

  const previousLastDay = supervisor.lastDay;

  const previousAssignedGroups = new Set(supervisor.assignedGroups);

  const previousAssignedProfessors = new Set(supervisor.assignedProfessors);

  const previousOccupiedSlots = new Set(supervisor.occupiedSlots);

  const previousByDay = {
    ...supervisor.byDay,
  };

  const previousByDayPeriods = {};

  for (const [day, periods] of Object.entries(supervisor.byDayPeriods || {})) {
    previousByDayPeriods[day] = new Set(periods);
  }

  const previousByDayPeriodRanks = {};

  for (const [day, ranks] of Object.entries(
    supervisor.byDayPeriodRanks || {},
  )) {
    previousByDayPeriodRanks[day] = new Set(ranks);
  }

  const previousBundleAssignments = new Map(bundleAssignments);

  for (const bundle of professorBundles) {
    const ok = assignBundleToSupervisor(
      supervisor,
      bundle,
      result,
      bundleAssignments,
    );

    if (!ok) {
      result.splice(previousResultLength);

      supervisor.total = previousTotal;

      supervisor.lastDay = previousLastDay;

      supervisor.assignedGroups = previousAssignedGroups;

      supervisor.assignedProfessors = previousAssignedProfessors;

      supervisor.occupiedSlots = previousOccupiedSlots;

      supervisor.byDay = previousByDay;

      supervisor.byDayPeriods = previousByDayPeriods;

      supervisor.byDayPeriodRanks = previousByDayPeriodRanks;

      bundleAssignments.clear();

      for (const [key, value] of previousBundleAssignments) {
        bundleAssignments.set(key, value);
      }

      return false;
    }
  }

  professorAssignments.set(professorKey, id);

  supervisor.assignedProfessors.add(professorKey);

  console.log(
    `✅ Professor assigned completely: ${professorKey} -> Supervisor ${id}`,
  );

  return true;
}

// ============================================================
// REBALANCE ENGINE
// ============================================================

function snapshotState(cand, result, bundleAssignments, professorAssignments) {
  const candSnapshot = {};

  for (const [id, supervisor] of Object.entries(cand)) {
    candSnapshot[id] = {
      id: supervisor.id,

      total: supervisor.total,

      lastDay: supervisor.lastDay,

      byDay: {
        ...supervisor.byDay,
      },

      byDayPeriods: Object.fromEntries(
        Object.entries(supervisor.byDayPeriods || {}).map(([day, periods]) => [
          day,
          new Set(periods),
        ]),
      ),

      byDayPeriodRanks: Object.fromEntries(
        Object.entries(supervisor.byDayPeriodRanks || {}).map(
          ([day, ranks]) => [day, new Set(ranks)],
        ),
      ),

      occupiedSlots: new Set(supervisor.occupiedSlots),

      assignedGroups: new Set(supervisor.assignedGroups),

      assignedProfessors: new Set(supervisor.assignedProfessors),
    };
  }

  return {
    cand: candSnapshot,

    result: result.map((row) => ({
      ...row,
    })),

    bundleAssignments: new Map(bundleAssignments),

    professorAssignments: new Map(professorAssignments),
  };
}

function restoreState(
  snapshot,
  cand,
  result,
  bundleAssignments,
  professorAssignments,
) {
  for (const [id, state] of Object.entries(snapshot.cand)) {
    if (!cand[id]) {
      continue;
    }

    cand[id].total = state.total;

    cand[id].lastDay = state.lastDay;

    cand[id].byDay = {
      ...state.byDay,
    };

    cand[id].byDayPeriods = Object.fromEntries(
      Object.entries(state.byDayPeriods).map(([day, periods]) => [
        day,
        new Set(periods),
      ]),
    );

    cand[id].byDayPeriodRanks = Object.fromEntries(
      Object.entries(state.byDayPeriodRanks).map(([day, ranks]) => [
        day,
        new Set(ranks),
      ]),
    );

    cand[id].occupiedSlots = new Set(state.occupiedSlots);

    cand[id].assignedGroups = new Set(state.assignedGroups);

    cand[id].assignedProfessors = new Set(state.assignedProfessors);
  }

  result.splice(
    0,
    result.length,
    ...snapshot.result.map((row) => ({
      ...row,
    })),
  );

  bundleAssignments.clear();

  for (const [key, value] of snapshot.bundleAssignments) {
    bundleAssignments.set(key, value);
  }

  professorAssignments.clear();

  for (const [key, value] of snapshot.professorAssignments) {
    professorAssignments.set(key, value);
  }
}

// ============================================================
// Rebuild Supervisor State
// ============================================================

function rebuildCandidateState(
  cand,
  bundles,
  professorGroups,
  bundleAssignments,
  professorAssignments,
  result,
) {
  for (const supervisor of Object.values(cand)) {
    supervisor.total = 0;

    supervisor.lastDay = null;

    supervisor.byDay = {};

    supervisor.byDayPeriods = {};

    supervisor.byDayPeriodRanks = {};

    supervisor.occupiedSlots = new Set();

    supervisor.assignedGroups = new Set();

    supervisor.assignedProfessors = new Set();
  }

  const bundleMap = new Map(bundles.map((bundle) => [bundle.key, bundle]));

  for (const [bundleKey, supervisorId] of bundleAssignments) {
    const supervisor = cand[Number(supervisorId)];

    const bundle = bundleMap.get(bundleKey);

    if (!supervisor || !bundle) {
      continue;
    }

    const representative = bundle.groups?.[0];

    if (!representative) {
      continue;
    }

    const day = dateISO(representative.date);

    const period = normalizePeriod(representative.period_label);

    const rank = getPeriodRank(period);

    if (!day || !period) {
      continue;
    }

    supervisor.total += 1;

    if (!supervisor.byDay[day]) {
      supervisor.byDay[day] = 0;
    }

    supervisor.byDay[day] += 1;

    if (!supervisor.byDayPeriods[day]) {
      supervisor.byDayPeriods[day] = new Set();
    }

    supervisor.byDayPeriods[day].add(period);

    if (!supervisor.byDayPeriodRanks[day]) {
      supervisor.byDayPeriodRanks[day] = new Set();
    }

    if (rank !== null && rank !== undefined) {
      supervisor.byDayPeriodRanks[day].add(rank);
    }

    supervisor.occupiedSlots.add(getSupervisorSlotKey(day, period));

    for (const group of bundle.groups || []) {
      const groupId = Number(group.id);

      if (Number.isFinite(groupId)) {
        supervisor.assignedGroups.add(groupId);
      }
    }
  }

  for (const professor of professorGroups) {
    const supervisorId = professorAssignments.get(professor.key);

    if (supervisorId === undefined) {
      continue;
    }

    const supervisor = cand[Number(supervisorId)];

    if (!supervisor) {
      continue;
    }

    supervisor.assignedProfessors.add(professor.key);
  }

  for (const supervisor of Object.values(cand)) {
    const days = Object.keys(supervisor.byDay).sort();

    supervisor.lastDay = days.length ? days[days.length - 1] : null;
  }

  const assignedGroupIds = new Set(
    result.map((row) => Number(row.session_group_id)),
  );

  for (const supervisor of Object.values(cand)) {
    for (const groupId of [...supervisor.assignedGroups]) {
      if (!assignedGroupIds.has(Number(groupId))) {
        supervisor.assignedGroups.delete(groupId);
      }
    }
  }
}

// ============================================================
// Rebalance Metrics
// ============================================================

function calculateRebalanceMetrics(
  cand,
  selectedSupervisorIds,
  minimumEnabled,
  minimumTarget,
) {
  const totals = selectedSupervisorIds.map((id) =>
    Number(cand[id]?.total || 0),
  );

  if (!totals.length) {
    return {
      minTotal: 0,
      maxTotal: 0,
      difference: 0,
      sumSquared: 0,
      supervisorsBelowMinimum: 0,
      totalMinimumDeficit: 0,
    };
  }

  const minTotal = Math.min(...totals);

  const maxTotal = Math.max(...totals);

  const difference = maxTotal - minTotal;

  const average = totals.reduce((sum, value) => sum + value, 0) / totals.length;

  const sumSquared = totals.reduce(
    (sum, value) => sum + Math.pow(value - average, 2),
    0,
  );

  let supervisorsBelowMinimum = 0;

  let totalMinimumDeficit = 0;

  if (minimumEnabled) {
    for (const id of selectedSupervisorIds) {
      const total = Number(cand[id]?.total || 0);

      if (total < minimumTarget) {
        supervisorsBelowMinimum++;

        totalMinimumDeficit += minimumTarget - total;
      }
    }
  }

  return {
    minTotal,
    maxTotal,
    difference,
    sumSquared,
    supervisorsBelowMinimum,
    totalMinimumDeficit,
  };
}

// ============================================================
// Rebalance Objective
// ============================================================

function compareRebalanceMetrics(a, b, minimumEnabled) {
  if (minimumEnabled) {
    if (a.supervisorsBelowMinimum !== b.supervisorsBelowMinimum) {
      return a.supervisorsBelowMinimum - b.supervisorsBelowMinimum;
    }

    if (a.totalMinimumDeficit !== b.totalMinimumDeficit) {
      return a.totalMinimumDeficit - b.totalMinimumDeficit;
    }
  }

  if (a.difference !== b.difference) {
    return a.difference - b.difference;
  }

  if (a.sumSquared !== b.sumSquared) {
    return a.sumSquared - b.sumSquared;
  }

  return 0;
}

// ============================================================
// Rebalance Engine
// ============================================================

function rebalanceAssignments({
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
}) {
  console.log("========================================");

  console.log("⚖️ STARTING REBALANCE ENGINE");

  console.log("========================================");

  let metrics = calculateRebalanceMetrics(
    cand,
    selectedSupervisorIds,
    minimumEnabled,
    minimumTarget,
  );

  console.log("📊 Before rebalance:", metrics);

  // ملاحظة: تم رفع عامل التكرار (من 2 إلى 4) لإعطاء محرك التوازن
  // مساحة أكبر للتقارب نحو أقل فرق ممكن بين المشرفين، خصوصًا في
  // الخطط الكبيرة (عدد أساتذة/مشرفين مرتفع).
  const maxIterations = Math.max(
    20,
    professorGroups.length * selectedSupervisorIds.length * 4,
  );

  let iterations = 0;

  let totalMoves = 0;

  while (iterations < maxIterations) {
    iterations++;

    let bestMove = null;

    let bestMetrics = metrics;

    const sourceSupervisors = [...selectedSupervisorIds]
      .map((id) => cand[id])
      .filter(Boolean)
      .sort((a, b) => Number(b.total || 0) - Number(a.total || 0));

    const targetSupervisors = [...selectedSupervisorIds]
      .map((id) => cand[id])
      .filter(Boolean)
      .sort((a, b) => Number(a.total || 0) - Number(b.total || 0));

    for (const sourceSupervisor of sourceSupervisors) {
      if (!sourceSupervisor) {
        continue;
      }

      const currentMax = Number(sourceSupervisor.total || 0);

      const currentMin = targetSupervisors.length
        ? Number(targetSupervisors[0].total || 0)
        : 0;

      if (currentMax <= currentMin) {
        continue;
      }

      const sourceProfessors = professorGroups.filter(
        (professor) =>
          Number(professorAssignments.get(professor.key)) ===
          Number(sourceSupervisor.id),
      );

      for (const professor of sourceProfessors) {
        const forced = forcedProfessorAssignments.get(professor.key);

        if (forced) {
          continue;
        }

        for (const targetSupervisor of targetSupervisors) {
          if (Number(targetSupervisor.id) === Number(sourceSupervisor.id)) {
            continue;
          }

          if (
            !canSupervisorTakeProfessor(targetSupervisor, professor.bundles)
          ) {
            continue;
          }

          const snapshot = snapshotState(
            cand,
            result,
            bundleAssignments,
            professorAssignments,
          );

          const professorGroupIds = new Set();

          const professorBundleKeys = new Set();

          for (const bundle of professor.bundles) {
            professorBundleKeys.add(bundle.key);

            for (const group of bundle.groups || []) {
              const groupId = Number(group.id);

              if (Number.isFinite(groupId)) {
                professorGroupIds.add(groupId);
              }
            }
          }

          for (let i = result.length - 1; i >= 0; i--) {
            if (professorGroupIds.has(Number(result[i].session_group_id))) {
              result.splice(i, 1);
            }
          }

          for (const bundleKey of professorBundleKeys) {
            if (
              Number(bundleAssignments.get(bundleKey)) ===
              Number(sourceSupervisor.id)
            ) {
              bundleAssignments.delete(bundleKey);
            }
          }

          professorAssignments.delete(professor.key);

          rebuildCandidateState(
            cand,
            bundles,
            professorGroups,
            bundleAssignments,
            professorAssignments,
            result,
          );

          const assigned = assignProfessorToSupervisor(
            professor.key,
            professor.bundles,
            targetSupervisor.id,
            result,
            cand,
            bundleAssignments,
            professorAssignments,
          );

          if (!assigned) {
            restoreState(
              snapshot,
              cand,
              result,
              bundleAssignments,
              professorAssignments,
            );

            continue;
          }

          rebuildCandidateState(
            cand,
            bundles,
            professorGroups,
            bundleAssignments,
            professorAssignments,
            result,
          );

          const candidateMetrics = calculateRebalanceMetrics(
            cand,
            selectedSupervisorIds,
            minimumEnabled,
            minimumTarget,
          );

          const comparison = compareRebalanceMetrics(
            candidateMetrics,
            metrics,
            minimumEnabled,
          );

          if (comparison < 0) {
            const continuityScore = getPeriodContinuityScore(
              targetSupervisor,
              professor.bundles,
            );

            bestMove = {
              professor,
              from: sourceSupervisor.id,
              to: targetSupervisor.id,
              candidateMetrics,
              continuityScore,
            };

            bestMetrics = candidateMetrics;
          }

          restoreState(
            snapshot,
            cand,
            result,
            bundleAssignments,
            professorAssignments,
          );
        }
      }
    }

    if (!bestMove) {
      break;
    }

    const professor = bestMove.professor;

    const sourceSupervisor = cand[Number(bestMove.from)];

    const targetSupervisor = cand[Number(bestMove.to)];

    if (!sourceSupervisor || !targetSupervisor) {
      break;
    }

    const professorGroupIds = new Set();

    const professorBundleKeys = new Set();

    for (const bundle of professor.bundles) {
      professorBundleKeys.add(bundle.key);

      for (const group of bundle.groups || []) {
        const groupId = Number(group.id);

        if (Number.isFinite(groupId)) {
          professorGroupIds.add(groupId);
        }
      }
    }

    for (let i = result.length - 1; i >= 0; i--) {
      if (professorGroupIds.has(Number(result[i].session_group_id))) {
        result.splice(i, 1);
      }
    }

    for (const bundleKey of professorBundleKeys) {
      if (
        Number(bundleAssignments.get(bundleKey)) === Number(sourceSupervisor.id)
      ) {
        bundleAssignments.delete(bundleKey);
      }
    }

    professorAssignments.delete(professor.key);

    rebuildCandidateState(
      cand,
      bundles,
      professorGroups,
      bundleAssignments,
      professorAssignments,
      result,
    );

    const moved = assignProfessorToSupervisor(
      professor.key,
      professor.bundles,
      targetSupervisor.id,
      result,
      cand,
      bundleAssignments,
      professorAssignments,
    );

    if (!moved) {
      console.warn("⚠️ Rebalance move failed unexpectedly.");

      break;
    }

    rebuildCandidateState(
      cand,
      bundles,
      professorGroups,
      bundleAssignments,
      professorAssignments,
      result,
    );

    metrics = calculateRebalanceMetrics(
      cand,
      selectedSupervisorIds,
      minimumEnabled,
      minimumTarget,
    );

    totalMoves++;

    console.log(
      `🔄 REBALANCE MOVE #${totalMoves}: ${professor.professor} : Supervisor ${bestMove.from} → Supervisor ${bestMove.to}`,
    );

    console.log("📊 New metrics:", metrics);
  }

  console.log("========================================");

  console.log("⚖️ REBALANCE FINISHED");

  console.log("========================================");

  console.log("🔄 Total rebalance moves:", totalMoves);

  console.log("📊 Final rebalance metrics:", metrics);

  return {
    iterations,
    totalMoves,
    metrics,
  };
}

// ============================================================
// MINIMUM REBALANCE ENGINE
// ============================================================

function minimumRebalanceAssignments({
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
}) {
  console.log("========================================");
  console.log("🎯 STARTING MINIMUM REBALANCE ENGINE");
  console.log("========================================");

  if (!minimumEnabled) {
    console.log(
      "🎯 Minimum Rebalance skipped because Minimum Rule is disabled.",
    );

    return {
      moves: 0,
      iterations: 0,
      finalMetrics: calculateRebalanceMetrics(
        cand,
        selectedSupervisorIds,
        false,
        minimumTarget,
      ),
      remainingDeficit: [],
    };
  }

  rebuildCandidateState(
    cand,
    bundles,
    professorGroups,
    bundleAssignments,
    professorAssignments,
    result,
  );

  let metrics = calculateRebalanceMetrics(
    cand,
    selectedSupervisorIds,
    minimumEnabled,
    minimumTarget,
  );

  console.log("🎯 Minimum target:", minimumTarget);
  console.log("📊 BEFORE Minimum Rebalance:", metrics);

  // ملاحظة: تم رفع عامل التكرار (من 3 إلى 5) لضمان قدرة المحرك
  // على الوصول للحد الأدنى لكل المشرفين مع الحفاظ على العدالة.
  const maxIterations = Math.max(
    20,
    professorGroups.length * selectedSupervisorIds.length * 5,
  );

  let iterations = 0;
  let totalMoves = 0;

  while (iterations < maxIterations) {
    iterations++;

    let bestMove = null;
    let bestMetrics = metrics;

    const deficitSupervisors = [...selectedSupervisorIds]
      .map((id) => cand[Number(id)])
      .filter(Boolean)
      .filter((supervisor) => Number(supervisor.total || 0) < minimumTarget)
      .sort((a, b) => Number(a.total || 0) - Number(b.total || 0));

    if (!deficitSupervisors.length) {
      console.log("✅ All supervisors reached the minimum.");

      break;
    }

    const sourceSupervisors = [...selectedSupervisorIds]
      .map((id) => cand[Number(id)])
      .filter(Boolean)
      .sort((a, b) => Number(b.total || 0) - Number(a.total || 0));

    for (const targetSupervisor of deficitSupervisors) {
      const targetTotal = Number(targetSupervisor.total || 0);

      const targetDeficit = minimumTarget - targetTotal;

      if (targetDeficit <= 0) {
        continue;
      }

      for (const sourceSupervisor of sourceSupervisors) {
        if (Number(sourceSupervisor.id) === Number(targetSupervisor.id)) {
          continue;
        }

        const sourceTotal = Number(sourceSupervisor.total || 0);

        if (sourceTotal <= targetTotal) {
          continue;
        }

        const sourceProfessors = professorGroups.filter(
          (professor) =>
            Number(professorAssignments.get(professor.key)) ===
            Number(sourceSupervisor.id),
        );

        if (!sourceProfessors.length) {
          continue;
        }

        const orderedProfessors = [...sourceProfessors].sort((a, b) => {
          const sizeA = a.bundles.length;

          const sizeB = b.bundles.length;

          const exactA = sizeA === targetDeficit ? 0 : 1;

          const exactB = sizeB === targetDeficit ? 0 : 1;

          if (exactA !== exactB) {
            return exactA - exactB;
          }

          const overshootA = Math.abs(sizeA - targetDeficit);

          const overshootB = Math.abs(sizeB - targetDeficit);

          if (overshootA !== overshootB) {
            return overshootA - overshootB;
          }

          return sizeA - sizeB;
        });

        for (const professor of orderedProfessors) {
          if (forcedProfessorAssignments.has(professor.key)) {
            continue;
          }

          const workload = professor.bundles.length;

          if (!workload) {
            continue;
          }

          const projectedSource = sourceTotal - workload;

          const projectedTarget = targetTotal + workload;

          if (
            projectedSource < targetTotal &&
            projectedTarget > minimumTarget
          ) {
            continue;
          }

          const snapshot = snapshotState(
            cand,
            result,
            bundleAssignments,
            professorAssignments,
          );

          const professorGroupIds = new Set();

          const professorBundleKeys = new Set();

          for (const bundle of professor.bundles) {
            professorBundleKeys.add(bundle.key);

            for (const group of bundle.groups || []) {
              const groupId = Number(group.id);

              if (Number.isFinite(groupId)) {
                professorGroupIds.add(groupId);
              }
            }
          }

          for (let i = result.length - 1; i >= 0; i--) {
            if (professorGroupIds.has(Number(result[i].session_group_id))) {
              result.splice(i, 1);
            }
          }

          for (const bundleKey of professorBundleKeys) {
            if (
              Number(bundleAssignments.get(bundleKey)) ===
              Number(sourceSupervisor.id)
            ) {
              bundleAssignments.delete(bundleKey);
            }
          }

          professorAssignments.delete(professor.key);

          rebuildCandidateState(
            cand,
            bundles,
            professorGroups,
            bundleAssignments,
            professorAssignments,
            result,
          );

          const freshTarget = cand[Number(targetSupervisor.id)];

          if (!freshTarget) {
            restoreState(
              snapshot,
              cand,
              result,
              bundleAssignments,
              professorAssignments,
            );

            continue;
          }

          const canTake = canSupervisorTakeProfessor(
            freshTarget,
            professor.bundles,
          );

          if (!canTake) {
            restoreState(
              snapshot,
              cand,
              result,
              bundleAssignments,
              professorAssignments,
            );

            continue;
          }

          const assigned = assignProfessorToSupervisor(
            professor.key,
            professor.bundles,
            freshTarget.id,
            result,
            cand,
            bundleAssignments,
            professorAssignments,
          );

          if (!assigned) {
            restoreState(
              snapshot,
              cand,
              result,
              bundleAssignments,
              professorAssignments,
            );

            continue;
          }

          rebuildCandidateState(
            cand,
            bundles,
            professorGroups,
            bundleAssignments,
            professorAssignments,
            result,
          );

          const candidateMetrics = calculateRebalanceMetrics(
            cand,
            selectedSupervisorIds,
            minimumEnabled,
            minimumTarget,
          );

          const comparison = compareRebalanceMetrics(
            candidateMetrics,
            metrics,
            minimumEnabled,
          );

          if (comparison < 0) {
            bestMove = {
              professor,
              from: sourceSupervisor.id,
              to: targetSupervisor.id,
              workload,
              candidateMetrics,
            };

            bestMetrics = candidateMetrics;
          }

          restoreState(
            snapshot,
            cand,
            result,
            bundleAssignments,
            professorAssignments,
          );
        }
      }
    }

    if (!bestMove) {
      console.log("⚠️ No valid Minimum Rebalance move found.");

      break;
    }

    const professor = bestMove.professor;

    const sourceSupervisor = cand[Number(bestMove.from)];

    const targetSupervisor = cand[Number(bestMove.to)];

    if (!sourceSupervisor || !targetSupervisor) {
      break;
    }

    const finalSnapshot = snapshotState(
      cand,
      result,
      bundleAssignments,
      professorAssignments,
    );

    const professorGroupIds = new Set();

    const professorBundleKeys = new Set();

    for (const bundle of professor.bundles) {
      professorBundleKeys.add(bundle.key);

      for (const group of bundle.groups || []) {
        const groupId = Number(group.id);

        if (Number.isFinite(groupId)) {
          professorGroupIds.add(groupId);
        }
      }
    }

    for (let i = result.length - 1; i >= 0; i--) {
      if (professorGroupIds.has(Number(result[i].session_group_id))) {
        result.splice(i, 1);
      }
    }

    for (const bundleKey of professorBundleKeys) {
      if (
        Number(bundleAssignments.get(bundleKey)) === Number(sourceSupervisor.id)
      ) {
        bundleAssignments.delete(bundleKey);
      }
    }

    professorAssignments.delete(professor.key);

    rebuildCandidateState(
      cand,
      bundles,
      professorGroups,
      bundleAssignments,
      professorAssignments,
      result,
    );

    const freshTarget = cand[Number(bestMove.to)];

    if (
      !freshTarget ||
      !canSupervisorTakeProfessor(freshTarget, professor.bundles)
    ) {
      console.warn("⚠️ Minimum Rebalance final validation failed. Restoring.");

      restoreState(
        finalSnapshot,
        cand,
        result,
        bundleAssignments,
        professorAssignments,
      );

      rebuildCandidateState(
        cand,
        bundles,
        professorGroups,
        bundleAssignments,
        professorAssignments,
        result,
      );

      break;
    }

    const moved = assignProfessorToSupervisor(
      professor.key,
      professor.bundles,
      freshTarget.id,
      result,
      cand,
      bundleAssignments,
      professorAssignments,
    );

    if (!moved) {
      console.warn("⚠️ Minimum Rebalance permanent move failed. Restoring.");

      restoreState(
        finalSnapshot,
        cand,
        result,
        bundleAssignments,
        professorAssignments,
      );

      rebuildCandidateState(
        cand,
        bundles,
        professorGroups,
        bundleAssignments,
        professorAssignments,
        result,
      );

      break;
    }

    rebuildCandidateState(
      cand,
      bundles,
      professorGroups,
      bundleAssignments,
      professorAssignments,
      result,
    );

    metrics = calculateRebalanceMetrics(
      cand,
      selectedSupervisorIds,
      minimumEnabled,
      minimumTarget,
    );

    totalMoves++;

    console.log(
      `🎯 MINIMUM MOVE #${totalMoves}: ${professor.professor_name || professor.key} : Supervisor ${bestMove.from} → Supervisor ${bestMove.to} (${bestMove.workload} periods)`,
    );

    console.log("📊 Minimum metrics:", metrics);
  }

  const remainingDeficit = selectedSupervisorIds
    .map((supervisorId) => {
      const supervisor = cand[Number(supervisorId)];

      const periods = Number(supervisor?.total || 0);

      return {
        supervisorId: Number(supervisorId),

        periods,

        minimum: minimumTarget,

        deficit: Math.max(0, minimumTarget - periods),
      };
    })
    .filter((item) => item.deficit > 0);

  console.log("========================================");
  console.log("🎯 MINIMUM REBALANCE FINISHED");
  console.log("========================================");

  console.log("🎯 Total minimum moves:", totalMoves);

  console.log("🎯 Minimum iterations:", iterations);

  console.log("📊 Final minimum metrics:", metrics);

  console.log("🎯 Remaining deficits:", remainingDeficit);

  return {
    moves: totalMoves,
    iterations,
    finalMetrics: metrics,
    remainingDeficit,
  };
}

// ============================================================
// Single Bundle Fit Check
// ============================================================

function canSingleBundleFitSupervisor(supervisor, day, period) {
  if (!supervisor || !day || !period) {
    return false;
  }

  const rank = getPeriodRank(period);

  if (!ALLOW_SAME_PERIOD_MULTIPLE_PROFESSORS) {
    const slotKey = getSupervisorSlotKey(day, period);

    if (supervisor.occupiedSlots && supervisor.occupiedSlots.has(slotKey)) {
      return false;
    }

    if (supervisor.byDayPeriods?.[day]?.has(period)) {
      return false;
    }

    if (
      rank !== null &&
      rank !== undefined &&
      supervisor.byDayPeriodRanks?.[day]?.has(rank)
    ) {
      return false;
    }
  }

  if (!isValidSequentialPeriodAttachment(supervisor, day, [rank])) {
    return false;
  }

  const previousDay = addDaysISO(day, -1);
  const nextDay = addDaysISO(day, 1);

  const existingDays = Object.keys(supervisor.byDay || {}).filter(
    (d) => Number(supervisor.byDay[d]) > 0,
  );

  if (existingDays.includes(previousDay) || existingDays.includes(nextDay)) {
    return false;
  }

  return true;
}

// ============================================================
// RELAXED FALLBACK ASSIGNMENT (LAST RESORT - LEVEL 1)
// ============================================================
//
// يتجاهل القواعد المرنة فقط:
//   - تسلسل الفترات
//   - منع الأيام المتتالية
//
// لكنه ما زال يشترط أن يأخذ مشرف واحد الأستاذ كاملًا.
// ============================================================

function assignRemainingProfessorsRelaxed({
  professorGroups,
  selectedSupervisorIds,
  cand,
  result,
  bundleAssignments,
  professorAssignments,
  forcedProfessorAssignments,
  conflicts,
  variant,
}) {
  console.log("========================================");
  console.log("🆘 STARTING RELAXED FALLBACK ASSIGNMENT");
  console.log("========================================");

  const unassigned = professorGroups.filter(
    (professor) => !professorAssignments.has(professor.key),
  );

  if (!unassigned.length) {
    console.log("🆘 No unassigned professors. Skipping relaxed fallback.");

    return { assigned: 0, stillUnassigned: 0, details: [] };
  }

  console.log(
    `🆘 Unassigned professors before relaxed pass: ${unassigned.length}`,
  );

  let assignedCount = 0;

  const details = [];

  for (const professor of unassigned) {
    if (forcedProfessorAssignments.has(professor.key)) {
      continue;
    }

    const ranked = [];

    for (const supervisorId of selectedSupervisorIds) {
      const supervisor = cand[Number(supervisorId)];

      if (!supervisor) continue;

      if (
        !canSupervisorTakeProfessor(supervisor, professor.bundles, {
          relaxed: true,
        })
      ) {
        continue;
      }

      ranked.push({
        supervisorId: Number(supervisorId),
        currentTotal: Number(supervisor.total || 0),
        tieBreaker: Number(supervisorId),
      });
    }

    if (!ranked.length) {
      // سيتم التعامل معه في مرحلة التغطية الإجبارية (Level 2)
      continue;
    }

    ranked.sort((a, b) => {
      if (a.currentTotal !== b.currentTotal) {
        return a.currentTotal - b.currentTotal;
      }

      return a.tieBreaker - b.tieBreaker;
    });

    const chosenSupervisorId = ranked[0].supervisorId;

    const ok = assignProfessorToSupervisor(
      professor.key,
      professor.bundles,
      chosenSupervisorId,
      result,
      cand,
      bundleAssignments,
      professorAssignments,
      { relaxed: true },
    );

    if (ok) {
      assignedCount++;

      for (let i = conflicts.length - 1; i >= 0; i--) {
        if (
          conflicts[i].type === "NO_SUPERVISOR_FOR_PROFESSOR" &&
          conflicts[i].professor_id === professor.professor_id &&
          conflicts[i].professor === professor.professor
        ) {
          conflicts.splice(i, 1);
        }
      }

      details.push({
        professor: professor.professor,
        professor_id: professor.professor_id,
        supervisor_id: chosenSupervisorId,
        periods: professor.bundles.length,
      });

      console.log(
        `🆘 RELAXED ASSIGNMENT: ${professor.professor} -> Supervisor ${chosenSupervisorId}`,
      );
    }
  }

  const stillUnassigned = unassigned.length - assignedCount;

  console.log("========================================");
  console.log("🆘 RELAXED FALLBACK ASSIGNMENT FINISHED");
  console.log("========================================");
  console.log(`🆘 Assigned via relaxed fallback: ${assignedCount}`);
  console.log(
    `🆘 Still unassigned (needs forced coverage): ${stillUnassigned}`,
  );

  return { assigned: assignedCount, stillUnassigned, details };
}

// ============================================================
// FORCE FULL COVERAGE (MANDATORY - LAST RESORT LEVEL 2)
// ============================================================
//
// الهدف:
// لا يُترك أي أستاذ بدون توزيع إطلاقًا.
//
// هنا نوزّع كل Bundle متبقٍ بشكل فردي، ويُسمح بتقسيم
// الأستاذ الواحد على أكثر من مشرف.
//
// القاعدة الوحيدة المتبقية هنا هي القاعدة الفيزيائية:
// المشرف لا يمكن أن يكون في نفس اليوم/الفترة مرتين
// (إلا إذا كان ALLOW_SAME_PERIOD_MULTIPLE_PROFESSORS مفعّلًا).
// ============================================================

function forceAssignAllRemainingBundles({
  professorGroups,
  selectedSupervisorIds,
  cand,
  result,
  bundleAssignments,
  professorAssignments,
  conflicts,
  splitProfessorKeys,
}) {
  console.log("========================================");
  console.log("🚨 STARTING FORCED FULL COVERAGE");
  console.log("========================================");

  let assignedBundles = 0;

  const details = [];

  for (const professor of professorGroups) {
    if (professorAssignments.has(professor.key)) {
      continue;
    }

    let professorAssignedSomething = false;

    for (const bundle of professor.bundles) {
      if (bundleAssignments.has(bundle.key)) {
        continue;
      }

      const representative = bundle.groups?.[0];

      if (!representative) {
        continue;
      }

      const day = dateISO(representative.date);

      const period = normalizePeriod(representative.period_label);

      if (!day || !period) {
        continue;
      }

      const slotKey = getSupervisorSlotKey(day, period);

      const rank = getPeriodRank(period);

      const candidates = selectedSupervisorIds
        .map((id) => cand[Number(id)])
        .filter(Boolean)
        .filter(
          (supervisor) =>
            ALLOW_SAME_PERIOD_MULTIPLE_PROFESSORS ||
            !supervisor.occupiedSlots.has(slotKey),
        )
        .sort((a, b) => {
          const totalA = Number(a.total || 0);
          const totalB = Number(b.total || 0);

          if (totalA !== totalB) {
            return totalA - totalB;
          }

          return Number(a.id) - Number(b.id);
        });

      if (!candidates.length) {
        conflicts.push({
          type: "PHYSICAL_SLOT_CONFLICT",

          professor: professor.professor,

          professor_id: professor.professor_id,

          date: day,

          period,

          message:
            "جميع المشرفين المختارين مشغولون فعليًا في نفس اليوم والفترة. يجب إضافة مشرف إضافي لتغطية هذه الفترة.",
        });

        continue;
      }

      const chosen = candidates[0];

      let attached = 0;

      for (const group of bundle.groups || []) {
        if (attachGroupToSupervisor(chosen, group, result)) {
          attached++;
        }
      }

      if (!attached) {
        continue;
      }

      if (!chosen.byDay[day]) {
        chosen.byDay[day] = 0;
      }

      chosen.byDay[day] += 1;

      chosen.total += 1;

      chosen.lastDay = day;

      if (!chosen.byDayPeriods[day]) {
        chosen.byDayPeriods[day] = new Set();
      }

      chosen.byDayPeriods[day].add(period);

      if (!chosen.byDayPeriodRanks[day]) {
        chosen.byDayPeriodRanks[day] = new Set();
      }

      if (rank !== null && rank !== undefined) {
        chosen.byDayPeriodRanks[day].add(rank);
      }

      chosen.occupiedSlots.add(slotKey);

      chosen.assignedProfessors.add(professor.key);

      bundleAssignments.set(bundle.key, chosen.id);

      splitProfessorKeys.add(professor.key);

      assignedBundles++;

      professorAssignedSomething = true;

      details.push({
        professor: professor.professor,
        professor_id: professor.professor_id,
        date: day,
        period,
        supervisor_id: chosen.id,
      });

      console.log(
        `🚨 FORCED COVERAGE: ${professor.professor} (${day} - ${period}) -> Supervisor ${chosen.id}`,
      );
    }

    if (professorAssignedSomething) {
      for (let i = conflicts.length - 1; i >= 0; i--) {
        if (
          conflicts[i].type === "NO_SUPERVISOR_FOR_PROFESSOR" &&
          conflicts[i].professor === professor.professor
        ) {
          conflicts.splice(i, 1);
        }
      }
    }
  }

  console.log("========================================");
  console.log("🚨 FORCED FULL COVERAGE FINISHED");
  console.log("========================================");
  console.log(`🚨 Forced coverage bundles: ${assignedBundles}`);

  return { assignedBundles, details };
}

// ============================================================
// SINGLE PROFESSOR DAY SPLIT (EQUAL DISTRIBUTION)
// ============================================================

function distributeSingleProfessorDays({
  professorGroups,
  forcedProfessorAssignments,
  selectedSupervisorIds,
  cand,
  result,
  bundleAssignments,
  conflicts,
  splitProfessorKeys,
  variant,
}) {
  console.log("========================================");
  console.log("📆 CHECKING SINGLE-PROFESSOR DAYS");
  console.log("========================================");

  if (selectedSupervisorIds.length <= 1) {
    console.log("📆 Skipped: only one supervisor selected.");
    return { daysSplit: 0 };
  }

  const dayMap = new Map();

  for (const professor of professorGroups) {
    for (const bundle of professor.bundles) {
      const representative = bundle.groups?.[0];
      if (!representative) continue;

      const day = dateISO(representative.date);
      if (!day) continue;

      if (!dayMap.has(day)) {
        dayMap.set(day, new Map());
      }

      const profMap = dayMap.get(day);

      if (!profMap.has(professor.key)) {
        profMap.set(professor.key, { professor, bundles: [] });
      }

      profMap.get(professor.key).bundles.push(bundle);
    }
  }

  let daysSplit = 0;

  for (const [day, profMap] of dayMap) {
    if (profMap.size !== 1) {
      continue;
    }

    const [[professorKey, entry]] = profMap;

    if (forcedProfessorAssignments.has(professorKey)) {
      continue;
    }

    const dayBundles = entry.bundles;

    if (dayBundles.length <= 1) {
      continue;
    }

    const professor = entry.professor;

    const sortedBundles = [...dayBundles].sort((a, b) =>
      sortPeriods(a.period, b.period),
    );

    console.log(
      `📆 Single-professor day detected: ${day} -> ${professor.professor} ` +
        `(${sortedBundles.length} periods) — splitting across ${selectedSupervisorIds.length} supervisors.`,
    );

    let splitCount = 0;

    for (const bundle of sortedBundles) {
      const representative = bundle.groups?.[0];
      if (!representative) continue;

      const bDay = dateISO(representative.date);
      const period = normalizePeriod(representative.period_label);

      const candidates = selectedSupervisorIds
        .map((id) => cand[Number(id)])
        .filter(Boolean)
        .filter((sup) => canSingleBundleFitSupervisor(sup, bDay, period))
        .sort((a, b) => {
          const totalA = Number(a.total || 0);
          const totalB = Number(b.total || 0);

          if (totalA !== totalB) {
            return totalA - totalB;
          }

          return (
            seededShuffle(`${variant}-${a.id}`) -
            seededShuffle(`${variant}-${b.id}`)
          );
        });

      if (!candidates.length) {
        conflicts.push({
          type: "NO_SUPERVISOR_FOR_SINGLE_PROFESSOR_DAY_PERIOD",
          professor: professor.professor,
          professor_id: professor.professor_id,
          date: bDay,
          period,
          message:
            "No supervisor could take this period while splitting a single-professor day.",
        });

        continue;
      }

      const chosen = candidates[0];

      const groupIds = [];

      for (const group of bundle.groups || []) {
        const attached = attachGroupToSupervisor(chosen, group, result);

        if (attached) {
          groupIds.push(Number(group.id));
        }
      }

      if (!groupIds.length) {
        continue;
      }

      if (!chosen.byDay[bDay]) {
        chosen.byDay[bDay] = 0;
      }

      chosen.byDay[bDay] += 1;
      chosen.total += 1;
      chosen.lastDay = bDay;

      if (!chosen.byDayPeriods[bDay]) {
        chosen.byDayPeriods[bDay] = new Set();
      }
      chosen.byDayPeriods[bDay].add(period);

      const rank = getPeriodRank(period);

      if (!chosen.byDayPeriodRanks[bDay]) {
        chosen.byDayPeriodRanks[bDay] = new Set();
      }

      if (rank !== null && rank !== undefined) {
        chosen.byDayPeriodRanks[bDay].add(rank);
      }

      chosen.occupiedSlots.add(getSupervisorSlotKey(bDay, period));
      chosen.assignedProfessors.add(professorKey);

      bundleAssignments.set(bundle.key, chosen.id);

      splitCount++;

      console.log(`   ↳ ${bDay} ${period} -> Supervisor ${chosen.id}`);
    }

    if (splitCount > 1) {
      splitProfessorKeys.add(professorKey);
      daysSplit++;
    }

    const dayBundleKeys = new Set(dayBundles.map((b) => b.key));

    professor.bundles = professor.bundles.filter(
      (b) => !dayBundleKeys.has(b.key),
    );
  }

  for (let i = professorGroups.length - 1; i >= 0; i--) {
    if (!professorGroups[i].bundles.length) {
      professorGroups.splice(i, 1);
    }
  }

  console.log("========================================");
  console.log(`📆 Single-professor days split: ${daysSplit}`);
  console.log("========================================");

  return { daysSplit };
}

// ============================================================
// MINIMUM SPLIT FALLBACK ENGINE (LAST RESORT)
// ============================================================

function minimumSplitFallback({
  cand,
  selectedSupervisorIds,
  professorGroups,
  bundles,
  bundleMap,
  result,
  bundleAssignments,
  professorAssignments,
  forcedProfessorAssignments,
  minimumEnabled,
  minimumTarget,
}) {
  console.log("=======================================");
  console.log("🧩 STARTING MINIMUM SPLIT FALLBACK");
  console.log("========================================");

  const details = [];
  const splitProfessorKeys = new Set();

  if (!minimumEnabled) {
    console.log("🧩 Split fallback skipped: minimum rule disabled.");

    return { moves: 0, details, splitProfessorKeys, remainingDeficit: [] };
  }

  function rebuild() {
    rebuildCandidateState(
      cand,
      bundles,
      professorGroups,
      bundleAssignments,
      professorAssignments,
      result,
    );
  }

  rebuild();

  // ملاحظة: رفعنا الحد الأدنى والمضاعف (20 → 40، ×2 → ×4) لإعطاء
  // محرك التقسيم مساحة كافية للوصول لأقرب توزيع عادل ممكن.
  const maxIterations = Math.max(40, bundles.length * 4);

  let iterations = 0;
  let totalMoves = 0;

  while (iterations < maxIterations) {
    iterations++;

    const deficitSupervisors = [...selectedSupervisorIds]
      .map((id) => cand[Number(id)])
      .filter(Boolean)
      .filter((s) => Number(s.total || 0) < minimumTarget)
      .sort((a, b) => Number(a.total || 0) - Number(b.total || 0));

    if (!deficitSupervisors.length) {
      console.log(
        "✅ All supervisors reached the minimum (after split fallback).",
      );

      break;
    }

    const target = deficitSupervisors[0];

    const donorSupervisors = [...selectedSupervisorIds]
      .map((id) => cand[Number(id)])
      .filter(Boolean)
      .filter(
        (s) =>
          Number(s.id) !== Number(target.id) &&
          Number(s.total || 0) - 1 >= minimumTarget,
      )
      .sort((a, b) => Number(b.total || 0) - Number(a.total || 0));

    let moved = false;

    for (const donor of donorSupervisors) {
      const donorBundleKeys = [...bundleAssignments.entries()]
        .filter(([, supId]) => Number(supId) === Number(donor.id))
        .map(([key]) => key);

      for (const bundleKey of donorBundleKeys) {
        const bundle = bundleMap.get(bundleKey);

        if (!bundle) {
          continue;
        }

        if (forcedProfessorAssignments.has(bundle.professorKey)) {
          continue;
        }

        const representative = bundle.groups?.[0];

        if (!representative) {
          continue;
        }

        const day = dateISO(representative.date);
        const period = normalizePeriod(representative.period_label);

        if (!canSingleBundleFitSupervisor(target, day, period)) {
          continue;
        }

        const groupIds = new Set(
          (bundle.groups || [])
            .map((g) => Number(g.id))
            .filter(Number.isFinite),
        );

        for (let i = result.length - 1; i >= 0; i--) {
          if (groupIds.has(Number(result[i].session_group_id))) {
            result.splice(i, 1);
          }
        }

        bundleAssignments.set(bundleKey, target.id);

        for (const group of bundle.groups || []) {
          result.push({
            session_group_id: Number(group.id),
            crn: group.crn,
            course_name: group.course_name ?? "",
            professor: group.professor_name || group.professor || "",
            room_number: group.room_number ?? null,
            professor_id: group.professor_id ?? null,
            date: dateISO(group.date),
            period: normalizePeriod(group.period_label),
            supervisor_id: target.id,
          });
        }

        rebuild();

        splitProfessorKeys.add(bundle.professorKey);

        details.push({
          professor: bundle.professor,
          professor_id: bundle.professor_id,
          date: day,
          period,
          from_supervisor_id: donor.id,
          to_supervisor_id: target.id,
        });

        totalMoves++;

        console.log(
          `🧩 SPLIT MOVE #${totalMoves}: ${bundle.professor} (${day} - ${period}) : Supervisor ${donor.id} → Supervisor ${target.id}`,
        );

        moved = true;

        break;
      }

      if (moved) {
        break;
      }
    }

    if (!moved) {
      console.log("⚠️ No valid split move found. Stopping split fallback.");

      break;
    }
  }

  const remainingDeficit = selectedSupervisorIds
    .map((id) => {
      const supervisor = cand[Number(id)];
      const periods = Number(supervisor?.total || 0);

      return {
        supervisorId: Number(id),
        periods,
        minimum: minimumTarget,
        deficit: Math.max(0, minimumTarget - periods),
      };
    })
    .filter((item) => item.deficit > 0);

  console.log("========================================");
  console.log("🧩 MINIMUM SPLIT FALLBACK FINISHED");
  console.log("========================================");
  console.log("🧩 Total split moves:", totalMoves);
  console.log("🧩 Split professors:", [...splitProfessorKeys]);
  console.log("🧩 Remaining deficit after split fallback:", remainingDeficit);

  return {
    moves: totalMoves,
    details,
    splitProfessorKeys,
    remainingDeficit,
  };
}

// ============================================================
// FINAL FAIRNESS SPLIT (BUNDLE-LEVEL, UNCONDITIONAL)
// ============================================================
//
// الهدف: تقليل الفرق بين أعلى وأقل مشرف حملاً إلى 0 أو 1 قدر
// الإمكان، حتى لو كانت قاعدة "الحد الأدنى" غير مفعّلة، وحتى لو
// كانت هناك دكاترة موزّعين مسبقًا عبر Relaxed/Forced Coverage
// خارج نطاق professorAssignments.
//
// يعمل على مستوى "محاضرة واحدة" (Bundle) وليس الدكتور كاملاً،
// وينقل من المشرف الأكثر حملاً إلى الأقل حملاً، بشرط:
// - عدم لمس الدكاترة الملزمين (Forced: Lock/Preassignment/Affinity)
// - عدم كسر القيد الفيزيائي للـ Slot (عبر canSingleBundleFitSupervisor)
// - عدم النزول بالمشرف المانح تحت الحد الأدنى إذا كان مفعّلاً
// ============================================================

function finalFairnessSplit({
  cand,
  selectedSupervisorIds,
  professorGroups,
  bundles,
  bundleMap,
  result,
  bundleAssignments,
  professorAssignments,
  forcedProfessorAssignments,
  minimumEnabled,
  minimumTarget,
  splitProfessorKeys,
}) {
  console.log("========================================");
  console.log("⚖️ STARTING FINAL FAIRNESS SPLIT");
  console.log("========================================");

  const details = [];

  if (selectedSupervisorIds.length <= 1) {
    console.log("⚖️ Skipped: only one supervisor selected.");
    return { moves: 0, details };
  }

  function rebuild() {
    rebuildCandidateState(
      cand,
      bundles,
      professorGroups,
      bundleAssignments,
      professorAssignments,
      result,
    );
  }

  rebuild();

  // ملاحظة: رفعنا الحد الأدنى والمضاعف (30 → 60، ×2 → ×4) لضمان
  // أن الجولة الأخيرة للعدالة تصل فعليًا إلى أقل فرق ممكن (0 أو 1)
  // بين المشرفين قبل التوقف، خصوصًا في الخطط الكبيرة.
  const maxIterations = Math.max(60, bundles.length * 4);

  let iterations = 0;
  let totalMoves = 0;

  while (iterations < maxIterations) {
    iterations++;

    const supervisorsList = selectedSupervisorIds
      .map((id) => cand[Number(id)])
      .filter(Boolean);

    if (!supervisorsList.length) break;

    const totals = supervisorsList.map((s) => Number(s.total || 0));

    const minTotal = Math.min(...totals);
    const maxTotal = Math.max(...totals);
    const difference = maxTotal - minTotal;

    if (difference <= 1) {
      console.log(
        `⚖️ Fairness target reached (difference=${difference}). Stopping.`,
      );
      break;
    }

    // الأقل حملاً أولاً (مستقبِلون محتملون)
    const targets = [...supervisorsList].sort(
      (a, b) => Number(a.total || 0) - Number(b.total || 0),
    );

    // الأكثر حملاً أولاً (مانحون محتملون)
    const donors = [...supervisorsList].sort(
      (a, b) => Number(b.total || 0) - Number(a.total || 0),
    );

    let moved = false;

    for (const target of targets) {
      const targetTotal = Number(target.total || 0);

      for (const donor of donors) {
        if (Number(donor.id) === Number(target.id)) continue;

        const donorTotal = Number(donor.total || 0);

        // النقل يجب أن يحسّن التوازن فعليًا
        if (donorTotal - 1 < targetTotal + 1) continue;

        // احترام الحد الأدنى للمانح إذا كانت القاعدة مفعّلة
        if (minimumEnabled && donorTotal - 1 < minimumTarget) continue;

        const donorBundleKeys = [...bundleAssignments.entries()]
          .filter(([, supId]) => Number(supId) === Number(donor.id))
          .map(([key]) => key);

        for (const bundleKey of donorBundleKeys) {
          const bundle = bundleMap.get(bundleKey);

          if (!bundle) continue;

          if (forcedProfessorAssignments.has(bundle.professorKey)) continue;

          const representative = bundle.groups?.[0];

          if (!representative) continue;

          const day = dateISO(representative.date);
          const period = normalizePeriod(representative.period_label);

          if (!canSingleBundleFitSupervisor(target, day, period)) continue;

          // ------------------------------------------------
          // تنفيذ النقل
          // ------------------------------------------------

          const groupIds = new Set(
            (bundle.groups || [])
              .map((g) => Number(g.id))
              .filter(Number.isFinite),
          );

          for (let i = result.length - 1; i >= 0; i--) {
            if (groupIds.has(Number(result[i].session_group_id))) {
              result.splice(i, 1);
            }
          }

          bundleAssignments.set(bundleKey, target.id);

          for (const group of bundle.groups || []) {
            result.push({
              session_group_id: Number(group.id),
              crn: group.crn,
              course_name: group.course_name ?? "",
              professor: group.professor_name || group.professor || "",
              room_number: group.room_number ?? null,
              professor_id: group.professor_id ?? null,
              date: dateISO(group.date),
              period: normalizePeriod(group.period_label),
              time_from: group.time_from ?? "",
              time_to: group.time_to ?? "",
              supervisor_id: target.id,
            });
          }

          rebuild();

          splitProfessorKeys.add(bundle.professorKey);

          details.push({
            professor: bundle.professor,
            professor_id: bundle.professor_id,
            date: day,
            period,
            from_supervisor_id: donor.id,
            to_supervisor_id: target.id,
          });

          totalMoves++;

          console.log(
            `⚖️ FAIRNESS MOVE #${totalMoves}: ${bundle.professor} (${day} - ${period}) : Supervisor ${donor.id} → Supervisor ${target.id}`,
          );

          moved = true;

          break;
        }

        if (moved) break;
      }

      if (moved) break;
    }

    if (!moved) {
      console.log("⚠️ No valid fairness move found. Stopping.");
      break;
    }
  }

  console.log("========================================");
  console.log("⚖️ FINAL FAIRNESS SPLIT FINISHED");
  console.log(`⚖️ Total fairness moves: ${totalMoves}`);
  console.log("========================================");

  return { moves: totalMoves, details };
}

// ============================================================
// SWAP REBALANCE ENGINE
// ============================================================

function swapProfessorsBetweenSupervisors({
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
}) {
  console.log("🔄 STARTING SWAP REBALANCE ENGINE");

  let metrics = calculateRebalanceMetrics(
    cand,
    selectedSupervisorIds,
    minimumEnabled,
    minimumTarget,
  );

  console.log("📊 Swap engine BEFORE:", metrics);

  let iterations = 0;
  let totalSwaps = 0;

  // ملاحظة: تم رفع عامل التكرار (من 2 إلى 4) لإعطاء محرك المبادلة
  // فرصة أكبر لإيجاد مبادلات تُحسّن العدالة بين المشرفين.
  const maxIterations = Math.max(
    20,
    professorGroups.length * selectedSupervisorIds.length * 4,
  );

  function getSupervisorProfessors(supervisorId) {
    return professorGroups.filter(
      (professor) =>
        Number(professorAssignments.get(professor.key)) ===
        Number(supervisorId),
    );
  }

  function removeProfessor(professor) {
    const professorGroupIds = new Set();

    const professorBundleKeys = new Set();

    for (const bundle of professor.bundles) {
      professorBundleKeys.add(bundle.key);

      for (const group of bundle.groups || []) {
        const groupId = Number(group.id);

        if (Number.isFinite(groupId)) {
          professorGroupIds.add(groupId);
        }
      }
    }

    for (let i = result.length - 1; i >= 0; i--) {
      if (professorGroupIds.has(Number(result[i].session_group_id))) {
        result.splice(i, 1);
      }
    }

    for (const bundleKey of professorBundleKeys) {
      bundleAssignments.delete(bundleKey);
    }

    professorAssignments.delete(professor.key);
  }

  while (iterations < maxIterations) {
    iterations++;

    let bestSwap = null;

    let bestMetrics = metrics;

    const supervisors = [...selectedSupervisorIds]
      .map((id) => cand[id])
      .filter(Boolean);

    const heavySupervisors = [...supervisors].sort(
      (a, b) => Number(b.total || 0) - Number(a.total || 0),
    );

    const lightSupervisors = [...supervisors].sort(
      (a, b) => Number(a.total || 0) - Number(b.total || 0),
    );

    for (const supervisorA of heavySupervisors) {
      for (const supervisorB of lightSupervisors) {
        if (Number(supervisorA.id) === Number(supervisorB.id)) {
          continue;
        }

        if (Number(supervisorA.total || 0) <= Number(supervisorB.total || 0)) {
          continue;
        }

        const professorsA = getSupervisorProfessors(supervisorA.id);

        const professorsB = getSupervisorProfessors(supervisorB.id);

        if (!professorsA.length || !professorsB.length) {
          continue;
        }

        for (const professorA of professorsA) {
          if (forcedProfessorAssignments.has(professorA.key)) {
            continue;
          }

          for (const professorB of professorsB) {
            if (forcedProfessorAssignments.has(professorB.key)) {
              continue;
            }

            const snapshot = snapshotState(
              cand,
              result,
              bundleAssignments,
              professorAssignments,
            );

            removeProfessor(professorA);

            removeProfessor(professorB);

            rebuildCandidateState(
              cand,
              bundles,
              professorGroups,
              bundleAssignments,
              professorAssignments,
              result,
            );

            const targetForA = cand[Number(supervisorB.id)];

            const targetForB = cand[Number(supervisorA.id)];

            if (!targetForA || !targetForB) {
              restoreState(
                snapshot,
                cand,
                result,
                bundleAssignments,
                professorAssignments,
              );

              continue;
            }

            const canA = canSupervisorTakeProfessor(
              targetForA,
              professorA.bundles,
            );

            if (!canA) {
              restoreState(
                snapshot,
                cand,
                result,
                bundleAssignments,
                professorAssignments,
              );

              continue;
            }

            const assignedA = assignProfessorToSupervisor(
              professorA.key,
              professorA.bundles,
              supervisorB.id,
              result,
              cand,
              bundleAssignments,
              professorAssignments,
            );

            if (!assignedA) {
              restoreState(
                snapshot,
                cand,
                result,
                bundleAssignments,
                professorAssignments,
              );

              continue;
            }

            const canB = canSupervisorTakeProfessor(
              targetForB,
              professorB.bundles,
            );

            if (!canB) {
              restoreState(
                snapshot,
                cand,
                result,
                bundleAssignments,
                professorAssignments,
              );

              continue;
            }

            const assignedB = assignProfessorToSupervisor(
              professorB.key,
              professorB.bundles,
              supervisorA.id,
              result,
              cand,
              bundleAssignments,
              professorAssignments,
            );

            if (!assignedB) {
              restoreState(
                snapshot,
                cand,
                result,
                bundleAssignments,
                professorAssignments,
              );

              continue;
            }

            rebuildCandidateState(
              cand,
              bundles,
              professorGroups,
              bundleAssignments,
              professorAssignments,
              result,
            );

            const candidateMetrics = calculateRebalanceMetrics(
              cand,
              selectedSupervisorIds,
              minimumEnabled,
              minimumTarget,
            );

            const comparison = compareRebalanceMetrics(
              candidateMetrics,
              metrics,
              minimumEnabled,
            );

            if (comparison < 0) {
              bestSwap = {
                professorA,
                professorB,

                fromA: supervisorA.id,

                fromB: supervisorB.id,

                toA: supervisorB.id,

                toB: supervisorA.id,

                candidateMetrics,
              };

              bestMetrics = candidateMetrics;
            }

            restoreState(
              snapshot,
              cand,
              result,
              bundleAssignments,
              professorAssignments,
            );
          }
        }
      }
    }

    if (!bestSwap) {
      break;
    }

    const professorA = bestSwap.professorA;

    const professorB = bestSwap.professorB;

    const finalSwapSnapshot = snapshotState(
      cand,
      result,
      bundleAssignments,
      professorAssignments,
    );

    removeProfessor(professorA);

    removeProfessor(professorB);

    rebuildCandidateState(
      cand,
      bundles,
      professorGroups,
      bundleAssignments,
      professorAssignments,
      result,
    );

    const assignedA = assignProfessorToSupervisor(
      professorA.key,
      professorA.bundles,
      bestSwap.toA,
      result,
      cand,
      bundleAssignments,
      professorAssignments,
    );

    if (!assignedA) {
      console.warn(
        "⚠️ Permanent swap failed while assigning Professor A. Restoring previous state.",
      );

      restoreState(
        finalSwapSnapshot,
        cand,
        result,
        bundleAssignments,
        professorAssignments,
      );

      rebuildCandidateState(
        cand,
        bundles,
        professorGroups,
        bundleAssignments,
        professorAssignments,
        result,
      );

      break;
    }

    const assignedB = assignProfessorToSupervisor(
      professorB.key,
      professorB.bundles,
      bestSwap.toB,
      result,
      cand,
      bundleAssignments,
      professorAssignments,
    );

    if (!assignedB) {
      console.warn(
        "⚠️ Permanent swap failed while assigning Professor B. Restoring previous state.",
      );

      restoreState(
        finalSwapSnapshot,
        cand,
        result,
        bundleAssignments,
        professorAssignments,
      );

      rebuildCandidateState(
        cand,
        bundles,
        professorGroups,
        bundleAssignments,
        professorAssignments,
        result,
      );

      break;
    }

    rebuildCandidateState(
      cand,
      bundles,
      professorGroups,
      bundleAssignments,
      professorAssignments,
      result,
    );

    metrics = calculateRebalanceMetrics(
      cand,
      selectedSupervisorIds,
      minimumEnabled,
      minimumTarget,
    );

    totalSwaps++;

    console.log(
      `🔄 SWAP #${totalSwaps}: ${professorA.professor} (${bestSwap.fromA} → ${bestSwap.toA}) ↔ ${professorB.professor} (${bestSwap.fromB} → ${bestSwap.toB})`,
    );

    console.log("📊 Swap metrics:", metrics);
  }

  console.log("========================================");

  console.log("🔄 SWAP REBALANCE FINISHED");

  console.log("========================================");

  console.log("🔄 Total swaps:", totalSwaps);

  console.log("🔄 Swap iterations:", iterations);

  console.log("📊 Final swap metrics:", metrics);

  return {
    iterations,
    totalSwaps,
    metrics,
  };
}

// ============================================================
// GENERATE PLAN
// ============================================================

async function generatePlan(
  planId,
  variant = 1,
  minimumPeriodsEnabled = false,
  minimumPeriods = 4,
) {
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

  console.log("========================================");
  console.log(`🚀 Generating plan ${planId}, variant ${variant}`);
  console.log("========================================");

  console.log("⚙️ Minimum periods enabled:", minimumEnabled);

  console.log("⚙️ Minimum periods target:", minimumTarget);

  console.log(
    "⚙️ Allow same period for multiple professors:",
    ALLOW_SAME_PERIOD_MULTIPLE_PROFESSORS,
  );

  // ==========================================================
  // Plan Context
  // ==========================================================

  const ctx = await getPlanContext(planId);

  if (!ctx) {
    throw new Error(`Plan ${planId} was not found.`);
  }

  const {
    groups = [],
    supervisors = [],
    pre = [],
    locks = [],
    aff = [],
    rooms = [],
  } = ctx;

  // =====================================================
  // خريطة أستاذ -> قاعة
  // =====================================================

  const professorRoomMap = new Map();

  for (const item of rooms) {
    const pid = Number(item.professor_id ?? item.professorId);

    const roomNumber = String(item.room_number ?? item.roomNumber ?? "").trim();

    if (Number.isFinite(pid) && roomNumber) {
      professorRoomMap.set(pid, roomNumber);
    }
  }

  for (const group of groups) {
    const pid = Number(group.professor_id);

    group.room_number = Number.isFinite(pid)
      ? (professorRoomMap.get(pid) ?? null)
      : null;
  }
  console.log("📦 Groups from getPlanContext:", groups.length);

  // ==========================================================
  // Validation
  // ==========================================================

  if (!groups.length) {
    console.warn("⚠️ No session groups found for this plan.");

    await saveAssignments(planId, []);

    return {
      success: true,
      planId,
      variant,

      assigned: 0,
      total: 0,

      assignedGroups: 0,
      totalGroups: 0,

      assignedBundles: 0,
      totalBundles: 0,

      conflicts: [],
      conflictsCount: 0,

      supervisorsUsed: 0,

      fairnessDifference: 0,

      professorUniquenessViolations: 0,

      rebalanceMoves: 0,
      rebalanceIterations: 0,

      minimumRebalanceMoves: 0,
      minimumRebalanceIterations: 0,
      minimumRebalanceRemainingDeficit: [],

      swapMoves: 0,
      swapIterations: 0,

      minimumPeriodsEnabled: minimumEnabled,
      minimumPeriods: minimumEnabled ? minimumTarget : null,

      minimumsReached: false,
      minimumStatistics: [],

      statistics: [],
      selectedSupervisors: [],
      assignments: [],
    };
  }

  if (!supervisors.length) {
    throw new Error("No supervisors are available for this plan.");
  }

  // ==========================================================
  // Duty Pool
  // ==========================================================

  let selectedSupervisorIds = [];

  try {
    const dutyPool = await getDutyPool(planId);

    if (Array.isArray(dutyPool) && dutyPool.length) {
      selectedSupervisorIds = dutyPool
        .map((item) =>
          Number(item.supervisor_id ?? item.supervisorId ?? item.id),
        )
        .filter(Number.isFinite);
    }
  } catch (error) {
    console.warn(
      "⚠️ Could not load duty pool. Falling back to active supervisors.",
    );
  }

  if (!selectedSupervisorIds.length) {
    selectedSupervisorIds = supervisors
      .map((s) => Number(s.id))
      .filter(Number.isFinite);
  }

  selectedSupervisorIds = Array.from(new Set(selectedSupervisorIds));

  console.log(
    `👥 Selected supervisors: ${selectedSupervisorIds.length}`,
    selectedSupervisorIds,
  );

  if (!selectedSupervisorIds.length) {
    throw new Error("No supervisors were selected for this plan.");
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

      occupiedSlots: new Set(),

      assignedGroups: new Set(),

      assignedProfessors: new Set(),
    };
  }

  // ==========================================================
  // Affinities
  // ==========================================================

  const professorAffinity = new Map();

  const professorNameAffinity = new Map();

  const result = [];

  const conflicts = [];

  const bundleAssignments = new Map();

  for (const item of aff) {
    const supervisorId = Number(item.supervisor_id ?? item.supervisorId);

    if (!Number.isFinite(supervisorId)) {
      console.warn("⚠️ Affinity skipped: invalid supervisor ID", item);

      continue;
    }

    if (!cand[supervisorId]) {
      conflicts.push({
        type: "AFFINITY_SUPERVISOR_NOT_SELECTED",

        professor: item.professor_name ?? item.name ?? "",

        professor_id: item.professor_id ?? null,

        supervisor_id: supervisorId,

        message: "The affinity supervisor is not in the selected Duty Pool.",
      });

      continue;
    }

    const professorId = Number(item.professor_id ?? item.professorId ?? null);

    if (Number.isFinite(professorId)) {
      professorAffinity.set(professorId, supervisorId);
    }

    const professorName = String(item.professor_name ?? item.name ?? "")
      .trim()
      .replace(/\s+/g, " ")
      .toLowerCase();

    if (professorName) {
      professorNameAffinity.set(professorName, supervisorId);
    }

    console.log(
      `🔗 Affinity loaded: "${professorName}" -> Supervisor ${supervisorId}`,
    );
  }

  // ==========================================================
  // Build Bundles
  // ==========================================================

  const bundlesMap = new Map();

  for (const group of groups) {
    const key = getBundleKey(group);

    if (!bundlesMap.has(key)) {
      bundlesMap.set(key, {
        key,

        date: dateISO(group.date),

        period: normalizePeriod(group.period_label),

        professor_id: group.professor_id ?? null,

        professor: group.professor_name ?? group.professor ?? "",

        professor_name: group.professor_name ?? group.professor ?? "",

        professorKey: getProfessorKey(group),

        professorNameKey: getProfessorNameKey(group),

        groups: [],
      });
    }

    bundlesMap.get(key).groups.push(group);
  }

  const bundles = Array.from(bundlesMap.values());

  console.log(`📦 Total CRNs / groups: ${groups.length}`);

  console.log(`📦 Total Professor + Date + Period bundles: ${bundles.length}`);

  // ==========================================================
  // Group By Professor
  // ==========================================================

  const professorsMap = new Map();

  for (const bundle of bundles) {
    const key = bundle.professorKey;

    if (!professorsMap.has(key)) {
      professorsMap.set(key, {
        key,

        professor_id: bundle.professor_id,

        professor: bundle.professor,

        professorNameKey: bundle.professorNameKey,

        bundles: [],
      });
    }

    professorsMap.get(key).bundles.push(bundle);
  }

  const professorGroups = Array.from(professorsMap.values());

  for (const professor of professorGroups) {
    professor.bundles.sort((a, b) => {
      const dateCompare = String(a.date).localeCompare(String(b.date));

      if (dateCompare !== 0) {
        return dateCompare;
      }

      return sortPeriods(a.period, b.period);
    });
  }

  professorGroups.sort((a, b) => {
    if (b.bundles.length !== a.bundles.length) {
      return b.bundles.length - a.bundles.length;
    }

    return String(a.professor).localeCompare(String(b.professor), "ar");
  });

  console.log(`👨‍🏫 Unique professors: ${professorGroups.length}`);

  const totalProfessorsBeforeSplit = professorGroups.length;

  // ==========================================================
  // Professor Assignments
  // ==========================================================

  const professorAssignments = new Map();

  const forcedProfessorAssignments = new Map();

  function registerForcedProfessor(group, supervisorId, source) {
    const professorKey = getProfessorKey(group);

    const id = Number(supervisorId);

    if (!Number.isFinite(id)) {
      return;
    }

    if (!cand[id]) {
      conflicts.push({
        type: "SUPERVISOR_NOT_SELECTED",

        professor: group.professor_name ?? group.professor ?? "",

        professor_id: group.professor_id ?? null,

        supervisor_id: id,

        source,

        message: "The required supervisor is not in the selected Duty Pool.",
      });

      return;
    }

    if (!forcedProfessorAssignments.has(professorKey)) {
      forcedProfessorAssignments.set(professorKey, {
        supervisorId: id,
        source,
      });

      return;
    }

    const existing = forcedProfessorAssignments.get(professorKey);

    if (Number(existing.supervisorId) === id) {
      return;
    }

    conflicts.push({
      type: "PROFESSOR_MULTIPLE_SUPERVISORS",

      professor: group.professor_name ?? group.professor ?? "",

      professor_id: group.professor_id ?? null,

      supervisor_1: existing.supervisorId,

      supervisor_2: id,

      message: `Professor is forced to two different supervisors by ${existing.source} and ${source}.`,
    });
  }

  // ==========================================================
  // Locks
  // ==========================================================

  for (const lock of locks) {
    const sessionGroupId = Number(lock.session_group_id ?? lock.sessionGroupId);

    const supervisorId = Number(lock.supervisor_id ?? lock.supervisorId);

    if (!Number.isFinite(sessionGroupId) || !Number.isFinite(supervisorId)) {
      continue;
    }

    const group = groups.find((g) => Number(g.id) === sessionGroupId);

    if (!group) {
      continue;
    }

    registerForcedProfessor(group, supervisorId, "lock");
  }

  // ==========================================================
  // Preassignments
  // ==========================================================

  for (const item of pre) {
    const sessionGroupId = Number(item.session_group_id ?? item.sessionGroupId);

    const supervisorId = Number(item.supervisor_id ?? item.supervisorId);

    if (!Number.isFinite(sessionGroupId) || !Number.isFinite(supervisorId)) {
      continue;
    }

    const group = groups.find((g) => Number(g.id) === sessionGroupId);

    if (!group) {
      continue;
    }

    registerForcedProfessor(group, supervisorId, "preassignment");
  }

  // ==========================================================
  // Apply Professor Affinities
  // ==========================================================

  for (const professor of professorGroups) {
    let affinitySupervisor = null;

    if (
      professor.professor_id !== null &&
      professor.professor_id !== undefined
    ) {
      affinitySupervisor = professorAffinity.get(
        Number(professor.professor_id),
      );
    }

    if (affinitySupervisor === null || affinitySupervisor === undefined) {
      affinitySupervisor = professorNameAffinity.get(
        professor.professorNameKey,
      );
    }

    if (affinitySupervisor !== null && affinitySupervisor !== undefined) {
      const supervisorId = Number(affinitySupervisor);

      console.log(
        `🎯 AFFINITY FOUND: Professor "${professor.professor}" -> Supervisor ${supervisorId}`,
      );

      if (!cand[supervisorId]) {
        conflicts.push({
          type: "AFFINITY_SUPERVISOR_NOT_SELECTED",

          professor: professor.professor,

          professor_id: professor.professor_id,

          supervisor_id: supervisorId,

          message:
            "The affinity supervisor is not part of the selected Duty Pool.",
        });

        continue;
      }

      const forced = forcedProfessorAssignments.get(professor.key);

      if (forced && Number(forced.supervisorId) !== supervisorId) {
        console.warn(
          `⚠️ Professor affinity conflicts with forced assignment: ${professor.professor}`,
        );

        conflicts.push({
          type: "AFFINITY_FORCED_CONFLICT",

          professor: professor.professor,

          professor_id: professor.professor_id,

          affinity_supervisor: supervisorId,

          forced_supervisor: forced.supervisorId,
        });

        continue;
      }

      forcedProfessorAssignments.set(professor.key, {
        supervisorId,
        source: "affinity",
      });

      console.log(
        `🔒 Professor "${professor.professor}" FORCED to Supervisor ${supervisorId}`,
      );
    }
  }

  // ==========================================================
  // Single-Professor-Day Split (Equal Distribution)
  // ==========================================================

  const singleDaySplitProfessorKeys = new Set();

  const singleDaySplitResult = distributeSingleProfessorDays({
    professorGroups,
    forcedProfessorAssignments,
    selectedSupervisorIds,
    cand,
    result,
    bundleAssignments,
    conflicts,
    splitProfessorKeys: singleDaySplitProfessorKeys,
    variant,
  });

  console.log("📆 Single-Professor-Day Split result:", singleDaySplitResult);

  // ==========================================================
  // Forced Professors FIRST
  // ==========================================================

  for (const professor of professorGroups) {
    const forced = forcedProfessorAssignments.get(professor.key);

    if (!forced) {
      continue;
    }

    const supervisor = cand[Number(forced.supervisorId)];

    if (!supervisor) {
      conflicts.push({
        type: "SUPERVISOR_NOT_FOUND",

        professor: professor.professor,

        supervisor_id: forced.supervisorId,
      });

      continue;
    }

    let ok = assignProfessorToSupervisor(
      professor.key,
      professor.bundles,
      forced.supervisorId,
      result,
      cand,
      bundleAssignments,
      professorAssignments,
    );

    // محاولة ثانية بوضع Relaxed حتى لا يضيع الإلزام
    if (!ok) {
      ok = assignProfessorToSupervisor(
        professor.key,
        professor.bundles,
        forced.supervisorId,
        result,
        cand,
        bundleAssignments,
        professorAssignments,
        { relaxed: true },
      );
    }

    if (!ok) {
      conflicts.push({
        type: "PROFESSOR_CANNOT_FIT_FORCED_SUPERVISOR",

        professor: professor.professor,

        professor_id: professor.professor_id,

        supervisor_id: forced.supervisorId,

        periods: professor.bundles.length,

        message:
          "The professor cannot fit completely into the forced supervisor because of period conflicts.",
      });
    }
  }

  // ==========================================================
  // Remaining Professors (Normal Fair Distribution)
  // ==========================================================

  console.log("========================================");
  console.log("⚖️ STARTING NORMAL FAIR DISTRIBUTION");
  console.log("========================================");

  for (const professor of professorGroups) {
    if (professorAssignments.has(professor.key)) {
      continue;
    }

    const allMinimumsReached = hasReachedAllMinimums(
      cand,
      selectedSupervisorIds,
      minimumEnabled,
      minimumTarget,
    );

    const ranked = [];

    for (const supervisorId of selectedSupervisorIds) {
      const supervisor = cand[supervisorId];

      if (!supervisor) {
        continue;
      }

      if (!canSupervisorTakeProfessor(supervisor, professor.bundles)) {
        continue;
      }

      const workload = professor.bundles.length;

      const currentTotal = Number(supervisor.total || 0);

      const projectedTotal = currentTotal + workload;

      const minimumInfo = getMinimumInfo(
        supervisor,
        workload,
        minimumEnabled,
        minimumTarget,
      );

      const consecutiveScore = getProfessorConsecutiveScore(
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

      const professorCount = supervisor.assignedProfessors?.size || 0;

      const noProfessorYet = professorCount === 0;

      let score =
        periodContinuityScore * 1000000 +
        continuityPriority * 1000 +
        (noProfessorYet ? 100000000 : 0) -
        projectedTotal * 10000 -
        dailyLoad * 2;

      if (minimumEnabled && minimumInfo.needsMinimum && !allMinimumsReached) {
        score += 1000000000 + minimumInfo.deficit * 1000000;

        if (minimumInfo.reachesMinimum) {
          score += 50000000;
        }
      }

      ranked.push({
        supervisor,

        supervisorId,

        currentTotal,

        projectedTotal,

        consecutiveScore,

        periodContinuityScore,

        continuityPriority,

        dailyLoad,

        professorCount,

        minimumEnabled,

        minimumTarget,

        needsMinimum: minimumInfo.needsMinimum,

        minimumDeficit: minimumInfo.deficit,

        projectedMinimumDeficit: minimumInfo.projectedDeficit,

        reachesMinimum: minimumInfo.reachesMinimum,

        score,
      });
    }

    if (!ranked.length) {
      conflicts.push({
        type: "NO_SUPERVISOR_FOR_PROFESSOR",

        professor: professor.professor,

        professor_id: professor.professor_id,

        periods: professor.bundles.length,

        message:
          "No single supervisor can take all periods of this professor without a same-period conflict.",
      });

      console.warn(
        "⚠️ No supervisor can take entire professor:",
        professor.professor,
      );

      continue;
    }

    ranked.sort((a, b) => {
      if (minimumEnabled) {
        const aNeedsMinimum = a.currentTotal < minimumTarget;

        const bNeedsMinimum = b.currentTotal < minimumTarget;

        if (aNeedsMinimum !== bNeedsMinimum) {
          return aNeedsMinimum ? -1 : 1;
        }

        if (aNeedsMinimum && bNeedsMinimum) {
          if (a.minimumDeficit !== b.minimumDeficit) {
            return b.minimumDeficit - a.minimumDeficit;
          }

          if (a.projectedMinimumDeficit !== b.projectedMinimumDeficit) {
            return a.projectedMinimumDeficit - b.projectedMinimumDeficit;
          }

          if (a.reachesMinimum !== b.reachesMinimum) {
            return a.reachesMinimum ? -1 : 1;
          }
        }
      }

      if (a.projectedTotal !== b.projectedTotal) {
        return a.projectedTotal - b.projectedTotal;
      }

      if (a.periodContinuityScore !== b.periodContinuityScore) {
        return b.periodContinuityScore - a.periodContinuityScore;
      }

      if (a.continuityPriority !== b.continuityPriority) {
        return b.continuityPriority - a.continuityPriority;
      }

      if (a.consecutiveScore !== b.consecutiveScore) {
        return b.consecutiveScore - a.consecutiveScore;
      }

      if (a.dailyLoad !== b.dailyLoad) {
        return a.dailyLoad - b.dailyLoad;
      }

      if (a.professorCount !== b.professorCount) {
        return a.professorCount - b.professorCount;
      }

      return (
        seededShuffle(`${variant}-${a.supervisorId}`) -
        seededShuffle(`${variant}-${b.supervisorId}`)
      );
    });

    const selected = ranked[0];

    const ok = assignProfessorToSupervisor(
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
        type: "PROFESSOR_ASSIGNMENT_FAILED",

        professor: professor.professor,

        professor_id: professor.professor_id,

        supervisor_id: selected.supervisorId,
      });

      console.warn(
        `⚠️ Failed assigning professor ${professor.professor} to supervisor ${selected.supervisorId}`,
      );
    }
  }

  // ==========================================================
  // RELAXED FALLBACK (LEVEL 1)
  // ==========================================================

  const relaxedFallbackResult = assignRemainingProfessorsRelaxed({
    professorGroups,
    selectedSupervisorIds,
    cand,
    result,
    bundleAssignments,
    professorAssignments,
    forcedProfessorAssignments,
    conflicts,
    variant,
  });

  console.log("🆘 Relaxed Fallback result:", relaxedFallbackResult);

  // ==========================================================
  // FORCED FULL COVERAGE (LEVEL 2) — التوزيع الإجباري
  // ==========================================================

  const forcedCoverageResult = forceAssignAllRemainingBundles({
    professorGroups,
    selectedSupervisorIds,
    cand,
    result,
    bundleAssignments,
    professorAssignments,
    conflicts,
    splitProfessorKeys: singleDaySplitProfessorKeys,
  });

  console.log("🚨 Forced Coverage result:", forcedCoverageResult);

  // ==========================================================
  // Rebalance Engine
  // ==========================================================

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

  // ==========================================================
  // Minimum Rebalance Engine
  // ==========================================================

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

  // ==========================================================
  // Swap Rebalance Engine
  // ==========================================================

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

  // ==========================================================
  // Minimum Split Fallback Engine (last resort)
  // ==========================================================

  const bundleMapForSplit = new Map(
    bundles.map((bundle) => [bundle.key, bundle]),
  );

  const minimumSplitResult = minimumSplitFallback({
    cand,
    selectedSupervisorIds,
    professorGroups,
    bundles,
    bundleMap: bundleMapForSplit,
    result,
    bundleAssignments,
    professorAssignments,
    forcedProfessorAssignments,
    minimumEnabled,
    minimumTarget,
  });

  console.log("🧩 Minimum Split Fallback result:", minimumSplitResult);

  const finalFairnessSplitProfessorKeys = new Set();

  const finalFairnessResult = finalFairnessSplit({
    cand,
    selectedSupervisorIds,
    professorGroups,
    bundles,
    bundleMap: bundleMapForSplit,
    result,
    bundleAssignments,
    professorAssignments,
    forcedProfessorAssignments,
    minimumEnabled,
    minimumTarget,
    splitProfessorKeys: finalFairnessSplitProfessorKeys,
  });

  console.log("⚖️ Final Fairness Split result:", finalFairnessResult);

  // ==========================================================
  // FINAL CONSISTENCY REBUILD
  // ==========================================================

  rebuildCandidateState(
    cand,
    bundles,
    professorGroups,
    bundleAssignments,
    professorAssignments,
    result,
  );

  // ==========================================================
  // COVERAGE CHECK
  // ==========================================================

  const unassignedBundles = bundles.filter(
    (bundle) => !bundleAssignments.has(bundle.key),
  );

  const uncoveredProfessors = [
    ...new Set(unassignedBundles.map((bundle) => bundle.professor)),
  ];

  console.log("========================================");
  console.log("🔍 FINAL COVERAGE CHECK");
  console.log("========================================");
  console.log(
    `📦 Bundles assigned: ${bundleAssignments.size}/${bundles.length}`,
  );
  console.log(`👨‍🏫 Professors in plan: ${totalProfessorsBeforeSplit}`);
  console.log(`⚠️ Uncovered professors: ${uncoveredProfessors.length}`);

  if (uncoveredProfessors.length) {
    console.warn("⚠️ Uncovered professors list:", uncoveredProfessors);
  }

  console.log("📦 result length:", result.length);

  // ==========================================================
  // Save
  // ==========================================================

  await saveAssignments(planId, result);

  console.log(`💾 Assignments saved for plan ${planId}`);

  // ==========================================================
  // Statistics
  // ==========================================================

  const assignedGroupIds = new Set(
    result
      .map((r) => Number(r.session_group_id ?? r.sessionGroupId ?? r.id))
      .filter(Number.isInteger),
  );

  const assignedGroups = assignedGroupIds.size;

  const totalGroups = groups.length;

  const supervisorsUsed = Object.values(cand).filter(
    (c) => Number(c.total || 0) > 0,
  );

  const totals = Object.values(cand).map((c) => Number(c.total || 0));

  const minTotal = totals.length ? Math.min(...totals) : 0;

  const maxTotal = totals.length ? Math.max(...totals) : 0;

  const fairnessDifference = maxTotal - minTotal;

  const totalBundles = bundles.length;

  const assignedBundles = bundleAssignments.size;

  // ==========================================================
  // Professor Uniqueness
  // ==========================================================

  const professorSupervisorCheck = new Map();

  for (const row of result) {
    const professorKey = getProfessorKey({
      professor_id: row.professor_id,

      professor_name: row.professor,
    });

    if (!professorSupervisorCheck.has(professorKey)) {
      professorSupervisorCheck.set(professorKey, Number(row.supervisor_id));
    } else {
      const existing = professorSupervisorCheck.get(professorKey);

      if (Number(existing) !== Number(row.supervisor_id)) {
        professorSupervisorCheck.set(professorKey, "MULTIPLE");
      }
    }
  }

  let professorUniquenessViolations = 0;

  let intentionalMinimumSplits = 0;

  for (const [professorKey, supervisorId] of professorSupervisorCheck) {
    const assigned = professorAssignments.get(professorKey);

        if (supervisorId === "MULTIPLE") {
      if (
        minimumSplitResult.splitProfessorKeys.has(professorKey) ||
        singleDaySplitProfessorKeys.has(professorKey) ||
        finalFairnessSplitProfessorKeys.has(professorKey)
      ) {
        intentionalMinimumSplits++;
        continue;
      }

      professorUniquenessViolations++;
      continue;
    }

    if (assigned !== undefined && Number(assigned) !== Number(supervisorId)) {
      professorUniquenessViolations++;
    }
  }

  // ==========================================================
  // Minimum Statistics
  // ==========================================================

  const minimumStatistics = Object.values(cand).map((c) => {
    const supervisor = supervisors.find((s) => Number(s.id) === Number(c.id));

    const reached = minimumEnabled ? Number(c.total) >= minimumTarget : null;

    return {
      "Supervisor ID": c.id,

      Supervisor: supervisor?.name ?? c.id,

      "Actual Periods": c.total,

      "Minimum Periods": minimumEnabled ? minimumTarget : "",

      "Minimum Reached": minimumEnabled ? (reached ? "Yes" : "No") : "",

      "Used Days": Object.keys(c.byDay).length,
    };
  });

  const minimumsReached = hasReachedAllMinimums(
    cand,
    selectedSupervisorIds,
    minimumEnabled,
    minimumTarget,
  );

  console.log(
    `✅ Plan generated: ${assignedGroups}/${totalGroups} groups assigned`,
  );

  console.log(`📦 Bundles assigned: ${assignedBundles}/${totalBundles}`);

  console.log(`👥 Supervisors used: ${supervisorsUsed.length}`);

  console.log(`⚠️ Conflicts: ${conflicts.length}`);

  console.log(`⚖️ Fairness difference: ${fairnessDifference}`);

  // ==========================================================
  // Export Excel
  // ==========================================================

  const exportDir = path.join(__dirname, "../exports");

  if (!fs.existsSync(exportDir)) {
    fs.mkdirSync(exportDir, {
      recursive: true,
    });
  }

  const distributionRows = result.map((row) => {
    const supervisor = supervisors.find(
      (s) => Number(s.id) === Number(row.supervisor_id),
    );

    return {
      "Course Name": row.course_name ?? "",
      CRN: row.crn ?? "",
      Professor: row.professor ?? row.professor_name ?? "",
          Room: row.room_number ?? "",

      Date: row.date ?? "",
      "Time From": row.time_from ?? "",
      "Time To": row.time_to ?? "",
      Period: row.period ?? row.period_label ?? "",
      Supervisor: supervisor?.name ?? row.supervisor_id ?? "",
    };
  });

  const conflictsRows = conflicts.map((item) => ({
    Date: item.date ?? "",

    Period: item.period ?? "",

    Professor: item.professor ?? "",

    "Professor ID": item.professor_id ?? "",

    Type: item.type ?? "",

    "Supervisor ID": item.supervisor_id ?? "",

    Message: item.message ?? "",
  }));

  const statisticsRows = Object.values(cand).map((c) => {
    const supervisor = supervisors.find((s) => Number(s.id) === Number(c.id));

    const minimumReached = minimumEnabled ? c.total >= minimumTarget : null;

    return {
      "Supervisor ID": c.id,

      Supervisor: supervisor?.name ?? c.id,

      "Total Periods": c.total,

      "Minimum Periods": minimumEnabled ? minimumTarget : "",

      "Minimum Reached": minimumEnabled ? (minimumReached ? "Yes" : "No") : "",

      "Used Days": Object.keys(c.byDay).length,
    };
  });

  const professorRows = professorGroups.map((professor) => {
    const supervisorId = professorAssignments.get(professor.key);

    const supervisor = supervisors.find(
      (s) => Number(s.id) === Number(supervisorId),
    );

    return {
      "Professor ID": professor.professor_id,

      Professor: professor.professor,

      "Total Periods": professor.bundles.length,

      Supervisor: supervisor?.name ?? supervisorId ?? "موزّع على أكثر من مشرف",
    };
  });

  // ==========================================================
  // Workbook
  // ==========================================================

  const workbook = xlsx.utils.book_new();

  const distributionSheet = xlsx.utils.json_to_sheet(distributionRows);

  const conflictsSheet = xlsx.utils.json_to_sheet(
    conflictsRows.length
      ? conflictsRows
      : [
          {
            Status: "No conflicts",
          },
        ],
  );

  const statisticsSheet = xlsx.utils.json_to_sheet(statisticsRows);

  const professorsSheet = xlsx.utils.json_to_sheet(professorRows);

  xlsx.utils.book_append_sheet(workbook, distributionSheet, "Distribution");

  xlsx.utils.book_append_sheet(workbook, conflictsSheet, "Conflicts");

  xlsx.utils.book_append_sheet(workbook, statisticsSheet, "Statistics");

  xlsx.utils.book_append_sheet(
    workbook,
    professorsSheet,
    "Professor Assignment",
  );

  // ==========================================================
  // Relaxed Assignments Sheet
  // (✅ نُقلت إلى هنا بعد إنشاء workbook — كانت تسبب
  //  ReferenceError وتوقف توليد الخطة بالكامل)
  // ==========================================================

  if (relaxedFallbackResult.details.length) {
    const relaxedRows = relaxedFallbackResult.details.map((item) => {
      const supervisor = supervisors.find(
        (s) => Number(s.id) === Number(item.supervisor_id),
      );

      return {
        Professor: item.professor,
        "Professor ID": item.professor_id ?? "",
        Periods: item.periods,
        Supervisor: supervisor?.name ?? item.supervisor_id,
        Note: "تم التوزيع بتجاهل قاعدة تسلسل الفترات و/أو منع الأيام المتتالية بسبب قلة عدد المشرفين",
      };
    });

    const relaxedSheet = xlsx.utils.json_to_sheet(relaxedRows);

    xlsx.utils.book_append_sheet(workbook, relaxedSheet, "Relaxed Assignments");
  }

  // ==========================================================
  // Forced Coverage Sheet
  // ==========================================================

  if (forcedCoverageResult.details.length) {
    const forcedRows = forcedCoverageResult.details.map((item) => {
      const supervisor = supervisors.find(
        (s) => Number(s.id) === Number(item.supervisor_id),
      );

      return {
        Professor: item.professor,
        "Professor ID": item.professor_id ?? "",
        Date: item.date,
        Period: item.period,
        Supervisor: supervisor?.name ?? item.supervisor_id,
        Note: "تغطية إجبارية: تم تقسيم الأستاذ على أكثر من مشرف لضمان عدم ترك أي فترة بدون إشراف",
      };
    });

    const forcedSheet = xlsx.utils.json_to_sheet(forcedRows);

    xlsx.utils.book_append_sheet(workbook, forcedSheet, "Forced Coverage");
  }

  // ==========================================================
  // Minimum Periods Sheet
  // ==========================================================

  if (minimumEnabled) {
    const minimumSheet = xlsx.utils.json_to_sheet(minimumStatistics);

    xlsx.utils.book_append_sheet(workbook, minimumSheet, "Minimum Periods");
  }

  // ==========================================================
  // Minimum Split (Fallback) Sheet
  // ==========================================================

  if (minimumEnabled && minimumSplitResult?.details?.length) {
    const splitRows = minimumSplitResult.details.map((item) => {
      const fromSupervisor = supervisors.find(
        (s) => Number(s.id) === Number(item.from_supervisor_id),
      );

      const toSupervisor = supervisors.find(
        (s) => Number(s.id) === Number(item.to_supervisor_id),
      );

      return {
        Professor: item.professor,
        "Professor ID": item.professor_id ?? "",
        Date: item.date,
        Period: item.period,
        "From Supervisor": fromSupervisor?.name ?? item.from_supervisor_id,
        "To Supervisor": toSupervisor?.name ?? item.to_supervisor_id,
        Note: "تقسيم اضطراري (ملاذ أخير) لتغطية الحد الأدنى فقط",
      };
    });

    const splitSheet = xlsx.utils.json_to_sheet(splitRows);

    xlsx.utils.book_append_sheet(
      workbook,
      splitSheet,
      "Minimum Split (Fallback)",
    );
  }

    if (finalFairnessResult?.details?.length) {
    const fairnessRows = finalFairnessResult.details.map((item) => {
      const fromSupervisor = supervisors.find(
        (s) => Number(s.id) === Number(item.from_supervisor_id),
      );

      const toSupervisor = supervisors.find(
        (s) => Number(s.id) === Number(item.to_supervisor_id),
      );

      return {
        Professor: item.professor,
        "Professor ID": item.professor_id ?? "",
        Date: item.date,
        Period: item.period,
        "From Supervisor": fromSupervisor?.name ?? item.from_supervisor_id,
        "To Supervisor": toSupervisor?.name ?? item.to_supervisor_id,
        Note: "نقل لتحقيق أقصى عدالة ممكنة بين المشرفين (Final Fairness Pass)",
      };
    });

    const fairnessSheet = xlsx.utils.json_to_sheet(fairnessRows);

    xlsx.utils.book_append_sheet(
      workbook,
      fairnessSheet,
      "Final Fairness Split",
    );
  }
  // ==========================================================
  // Export Path
  // ==========================================================

  const exportPath = path.join(exportDir, `plan_${planId}.xlsx`);

  xlsx.writeFile(workbook, exportPath);

  console.log(`📄 Excel exported: ${exportPath}`);

  // ==========================================================
  // Return
  // ==========================================================

  return {
    success: true,

    planId,

    variant,

    assigned: assignedGroups,

    total: totalGroups,

    assignedGroups,
    totalGroups,

    assignedBundles,
    totalBundles,

    conflicts,

    conflictsCount: conflicts.length,

    supervisorsUsed: supervisorsUsed.length,

    fairnessDifference,

    idealFairnessRange: "0-1",

    professorUniquenessViolations,

    rebalanceMoves: rebalanceResult?.moves ?? rebalanceResult?.totalMoves ?? 0,

    rebalanceIterations: rebalanceResult?.iterations ?? 0,

    minimumRebalanceMoves: minimumRebalanceResult?.moves ?? 0,

    minimumRebalanceIterations: minimumRebalanceResult?.iterations ?? 0,

    minimumRebalanceRemainingDeficit:
      minimumRebalanceResult?.remainingDeficit ?? [],

    swapMoves:
      swapRebalanceResult?.swaps ?? swapRebalanceResult?.totalSwaps ?? 0,

    swapIterations: swapRebalanceResult?.iterations ?? 0,

    minimumSplitMoves: minimumSplitResult?.moves ?? 0,

    minimumSplitDetails: minimumSplitResult?.details ?? [],

    finalFairnessMoves: finalFairnessResult?.moves ?? 0,

    finalFairnessDetails: finalFairnessResult?.details ?? [],
    relaxedAssignmentsCount: relaxedFallbackResult.assigned,
    relaxedAssignmentDetails: relaxedFallbackResult.details,

    forcedCoverageBundles: forcedCoverageResult.assignedBundles,
    forcedCoverageDetails: forcedCoverageResult.details,

    uncoveredProfessors,
    uncoveredProfessorsCount: uncoveredProfessors.length,

    totalProfessors: totalProfessorsBeforeSplit,

    intentionalMinimumSplits,

    minimumRemainingDeficitAfterSplit:
      minimumSplitResult?.remainingDeficit ?? [],

    exportPath,

    statistics: statisticsRows,

    minimumPeriodsEnabled: minimumEnabled,

    minimumPeriods: minimumEnabled ? minimumTarget : null,

    minimumsReached,

    minimumStatistics,

    selectedSupervisors: selectedSupervisorIds,

    assignments: result,
  };
}

// ============================================================
// Daily Pools
// ============================================================

function buildDailyPoolsForPlan(days, supervisorIds) {
  const pools = {};

  const ids = supervisorIds.map(Number).filter(Number.isFinite);

  if (!ids.length) {
    return pools;
  }

  let previousPool = [];

  for (let i = 0; i < days.length; i++) {
    const day = days[i];

    let available = ids.filter((id) => !previousPool.includes(id));

    if (!available.length) {
      available = [...ids];
    }

    const shuffled = seededShuffle(available, i + 1);

    pools[day] = shuffled.length ? shuffled : [...ids];

    previousPool = pools[day];
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