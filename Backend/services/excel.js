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

  // الهيدر موجود بالصف الثالث
  const rawRows = xlsx.utils.sheet_to_json(sheet, {
    range: 2,
    defval: null,
    raw: false,
  });

  // ============================================================
  // DEBUG: بيانات 2025-11-06
  // ============================================================

  const debugNov6 = rawRows.filter((row) => {
    const dateValue =
      row["التاريخ ميلادي"] ??
      row["Date"] ??
      row["DATE"] ??
      row["date"] ??
      null;

    return dateISO(dateValue) === "2025-11-06";
  });

  console.log("========================================");
  console.log("🔎 EXCEL 2025-11-06 DEBUG");
  console.log("========================================");
  console.log("📊 Rows for 2025-11-06:", debugNov6.length);

  const debugCrns = debugNov6.map((row) =>
    String(row["CRN"] ?? row["crn"] ?? "").trim(),
  );

  const duplicateCrns = debugCrns.filter(
    (crn, index) => crn && debugCrns.indexOf(crn) !== index,
  );

  console.log("🔢 Unique CRNs:", new Set(debugCrns).size);

  console.log("🔁 Duplicate CRNs:", [...new Set(duplicateCrns)]);

  console.table(
    debugNov6.map((row, index) => {
      const dateValue =
        row["التاريخ ميلادي"] ??
        row["Date"] ??
        row["DATE"] ??
        row["date"] ??
        "";

      return {
        excel_index: index + 1,

        CRN: row["CRN"] ?? row["crn"] ?? "",

        Professor:
          row["Prof"] ??
          row["Professor"] ??
          row["Professor Name"] ??
          row["professor_name"] ??
          row["professor"] ??
          "",

        OriginalDate: dateValue,

        NormalizedDate: dateISO(dateValue),

        Period:
          row["الفترة"] ??
          row["Period"] ??
          row["period_label"] ??
          row["period"] ??
          "",
      };
    }),
  );

  console.log("========================================");

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

  // ============================================================
  // Empty Excel
  // ============================================================

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

      // كل Group بده مشرف واحد
      required_supervisors: 1,

      // البيانات الأصلية
      original: row,

      // رقم صف Excel الحقيقي
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

  // إذا كانت String
  if (typeof value === "string") {
    const str = value.trim();

    if (!str) return "";

    // YYYY-MM-DD
    // أو YYYY-MM-DDTHH:mm:ss
    const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);

    if (isoMatch) {
      return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
    }

    // DD/MM/YYYY
    // أو DD-MM-YYYY
    const dateMatch = str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);

    if (dateMatch) {
      const day = String(dateMatch[1]).padStart(2, "0");
      const month = String(dateMatch[2]).padStart(2, "0");
      const year = dateMatch[3];

      return `${year}-${month}-${day}`;
    }

    // نجرب Date كحل أخير
    const d = new Date(str);

    if (!Number.isNaN(d.getTime())) {
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, "0");
      const day = String(d.getDate()).padStart(2, "0");

      return `${year}-${month}-${day}`;
    }
  }

  // Excel cellDates = true
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      return "";
    }

    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  }

  // أي قيمة أخرى
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

  const newDays = new Set(
    professorBundles.map((bundle) => dateISO(bundle.date)).filter(Boolean),
  );

  const newDaysArray = [...newDays];

  // الدكتور نفسه عنده أيام متتالية
  for (const day of newDaysArray) {
    const previousDay = addDaysISO(day, -1);
    const nextDay = addDaysISO(day, 1);

    if (newDays.has(previousDay) || newDays.has(nextDay)) {
      return true;
    }
  }

  // الأيام الجديدة متتالية مع أيام المشرف
  for (const day of newDaysArray) {
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

// Bundle = دكتور + تاريخ + فترة
// كل CRNs لنفس الدكتور والتاريخ والفترة بتظل مع بعض
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

  // ما نضيف نفس الـ Group مرتين
  if (supervisor.assignedGroups.has(groupId)) {
    return false;
  }

  supervisor.assignedGroups.add(groupId);

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

  // إذا الـ Bundle موزع من قبل
  if (bundleAssignments.has(bundle.key)) {
    const assignedSupervisorId = bundleAssignments.get(bundle.key);

    if (Number(assignedSupervisorId) !== Number(supervisor.id)) {
      return false;
    }

    return true;
  }

  // نفس المشرف ممنوع يأخذ نفس اليوم والفترة
  if (supervisor.occupiedSlots && supervisor.occupiedSlots.has(slotKey)) {
    return false;
  }

  // Defensive check
  if (!supervisor.byDayPeriods[day]) {
    supervisor.byDayPeriods[day] = new Set();
  }

  if (supervisor.byDayPeriods[day].has(period)) {
    return false;
  }

  if (!supervisor.byDayPeriodRanks[day]) {
    supervisor.byDayPeriodRanks[day] = new Set();
  }

  if (
    periodRank !== null &&
    periodRank !== undefined &&
    supervisor.byDayPeriodRanks[day].has(periodRank)
  ) {
    return false;
  }

  // نتأكد إن كل الـ Groups صالح
  for (const group of groups) {
    if (!group || group.id === null || group.id === undefined) {
      return false;
    }

    const groupId = Number(group.id);

    if (supervisor.assignedGroups.has(groupId)) {
      return false;
    }
  }

  // نضيف كل الـ Groups
  for (const group of groups) {
    const added = attachGroupToSupervisor(supervisor, group, result);

    if (!added) {
      return false;
    }
  }

  // نحدث بيانات المشرف
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

function canSupervisorTakeProfessor(supervisor, professorBundles) {
  if (!supervisor || !professorBundles?.length) {
    return false;
  }

  // الدكتور نفسه ما يكون عنده نفس اليوم + نفس الفترة مرتين
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

  // نفس المشرف ممنوع نفس اليوم + الفترة
  for (const bundle of professorBundles) {
    const representative = bundle.groups?.[0];

    const day = dateISO(representative.date);

    const period = normalizePeriod(representative.period_label);

    const slotKey = getSupervisorSlotKey(day, period);

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

  // ممنوع يومين متتاليين
  if (hasConsecutiveDayConflict(supervisor, professorBundles)) {
    return false;
  }

  return true;
}

// ============================================================
// Consecutive Period Score
// ============================================================

// بنعطي أفضلية للفترات اللي بتكون جنب بعض
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
      // المشرف عنده الفترة السابقة
      if (existingRanks.has(rank - 1)) {
        score += 100;
      }

      // المشرف عنده الفترة التالية
      if (existingRanks.has(rank + 1)) {
        score += 60;
      }

      // الفترات نفسها عند الدكتور
      if (candidateRanks.has(rank - 1)) {
        score += 50;
      }

      if (candidateRanks.has(rank + 1)) {
        score += 30;
      }

      // إذا الفترة بتكمل فراغ بين فترتين
      if (existingRanks.size > 1) {
        const sorted = [...existingRanks].sort((a, b) => a - b);

        const minRank = sorted[0];

        const maxRank = sorted[sorted.length - 1];

        if (rank > minRank && rank < maxRank && !existingRanks.has(rank)) {
          score += 80;
        }
      }

      // إذا عنده سلسلة فترات متتابعة
      if (existingRanks.has(rank - 1) && existingRanks.has(rank - 2)) {
        score += 40;
      }
    }
  }

  return score;
}

// ============================================================
// ⭐ Period Continuity Score - UPDATED
// ============================================================
//
// الهدف:
//
// نعطي الأولوية للمشرف الذي:
//
// 1. يكمل فترة بجانب فترة موجودة.
// 2. يغلق Gap موجود.
// 3. يزيد طول السلسلة المتصلة.
// 4. لا يفتح Gap جديد.
// 5. يحافظ قدر الإمكان على الفترات متجاورة.
//
// مثال:
//
// الموجود:
// 1,2
//
// المرشح:
// 3
//
// يصبح:
// 1,2,3
//
// => أولوية قوية جداً
//
// ------------------------------------------------------------
//
// الموجود:
//
// 1,2,4
//
// المرشح:
//
// 3
//
// يصبح:
//
// 1,2,3,4
//
// => أقوى حالة لأنها أغلقت Gap.
//
// ------------------------------------------------------------
//
// الموجود:
//
// 1,2
//
// المرشح:
//
// 4
//
// يصبح:
//
// 1,2,4
//
// => لا نريد هذا إذا كان هناك مشرف آخر يستطيع أخذ 4
// بدون تكوين Gap.
//
// ============================================================

function getPeriodContinuityScore(supervisor, professorBundles) {
  let score = 0;

  const candidateByDay = new Map();

  // ----------------------------------------------------------
  // نجمع فترات الدكتور حسب اليوم
  // ----------------------------------------------------------

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

  // ----------------------------------------------------------
  // Helper: أطول سلسلة متصلة
  // ----------------------------------------------------------

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

  // ----------------------------------------------------------
  // Helper: حساب الـ gaps الداخلية
  // ----------------------------------------------------------

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

  // ----------------------------------------------------------
  // Helper:
  // هل المرشح يضيف فترة متجاورة؟
  // ----------------------------------------------------------

  function hasAdjacentPeriod(existingRanks, rank) {
    return existingRanks.has(rank - 1) || existingRanks.has(rank + 1);
  }

  // ----------------------------------------------------------
  // لكل يوم
  // ----------------------------------------------------------

  for (const [day, candidateRanks] of candidateByDay) {
    const existingRanks = new Set(supervisor.byDayPeriodRanks?.[day] || []);

    // --------------------------------------------------------
    // الوضع قبل الإضافة
    // --------------------------------------------------------

    const beforeRun = longestContinuousRun(existingRanks);

    const beforeGaps = countInternalGaps(existingRanks);

    // --------------------------------------------------------
    // الوضع بعد الإضافة
    // --------------------------------------------------------

    const afterRanks = new Set(existingRanks);

    for (const rank of candidateRanks) {
      afterRanks.add(rank);
    }

    const afterRun = longestContinuousRun(afterRanks);

    const afterGaps = countInternalGaps(afterRanks);

    // ========================================================
    // 1. إغلاق Gap كامل
    // ========================================================
    //
    // مثال:
    //
    // 1,2,4 + 3
    //
    // هذه أقوى حالة.
    //
    // ========================================================

    for (const rank of candidateRanks) {
      const fillsGap =
        existingRanks.has(rank - 1) && existingRanks.has(rank + 1);

      if (fillsGap) {
        score += 1000000;
      }
    }

    // ========================================================
    // 2. تقليل عدد الـ Gaps
    // ========================================================

    const gapReduction = beforeGaps - afterGaps;

    if (gapReduction > 0) {
      score += gapReduction * 200000;
    }

    // ========================================================
    // 3. زيادة أطول سلسلة متصلة
    // ========================================================

    const runGrowth = afterRun - beforeRun;

    if (runGrowth > 0) {
      score += runGrowth * 100000;
    }

    // ========================================================
    // 4. إكمال الفترة السابقة مباشرة
    //
    // 1,2 + 3
    //
    // ========================================================

    for (const rank of candidateRanks) {
      if (existingRanks.has(rank - 1)) {
        score += 30000;
      }
    }

    // ========================================================
    // 5. إكمال الفترة التالية مباشرة
    //
    // 3 + 4
    //
    // ========================================================

    for (const rank of candidateRanks) {
      if (existingRanks.has(rank + 1)) {
        score += 25000;
      }
    }

    // ========================================================
    // 6. ربط سلسلتين
    //
    // مثال:
    //
    // 1,2 + 4,5
    //
    // والمرشح 3
    //
    // يصبح:
    //
    // 1,2,3,4,5
    //
    // ========================================================

    for (const rank of candidateRanks) {
      const connectsBothSides =
        existingRanks.has(rank - 1) && existingRanks.has(rank + 1);

      if (connectsBothSides) {
        score += 500000;
      }
    }

    // ========================================================
    // 7. إكمال سلسلة طويلة
    //
    // إذا المشرف عنده:
    //
    // 1,2
    //
    // والمرشح 3
    //
    // نعطي bonus إضافي.
    // ========================================================

    for (const rank of candidateRanks) {
      if (existingRanks.has(rank - 1) && existingRanks.has(rank - 2)) {
        score += 20000;
      }

      if (existingRanks.has(rank + 1) && existingRanks.has(rank + 2)) {
        score += 18000;
      }
    }

    // ========================================================
    // 8. إذا المرشح نفسه يحتوي على فترات متتابعة
    // ========================================================

    const candidateRun = longestContinuousRun(candidateRanks);

    if (candidateRun > 1) {
      score += (candidateRun - 1) * 10000;
    }

    // ========================================================
    // 9. Bonus عام لأي فترة بجانب فترة موجودة
    // ========================================================

    for (const rank of candidateRanks) {
      if (hasAdjacentPeriod(existingRanks, rank)) {
        score += 5000;
      }
    }

    // ========================================================
    // 10. عقوبة تكوين Gap جديد
    //
    // مثال:
    //
    // الموجود 1,2
    // المرشح 4
    //
    // النتيجة 1,2,4
    //
    // Gap جديد = 1
    //
    // ========================================================

    const newGaps = afterGaps - beforeGaps;

    if (newGaps > 0) {
      score -= newGaps * 100000;
    }

    // ========================================================
    // 11. إذا أصبحت كل الفترات داخل نطاق متصل
    // ========================================================

    if (afterRanks.size > 0 && afterGaps === 0) {
      score += 50000;
    }
  }

  return score;
}

// ============================================================
// ⭐ Continuity Priority
// ============================================================
//
// نحافظ على الدالة السابقة حتى ما نكسر أي منطق آخر.
//
// لكنها الآن تعتمد على الـ Period Continuity الأقوى.
//
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

  // ----------------------------------------------------------
  // أولاً: نتأكد أن الدكتور كامل ممكن يركب عند المشرف
  // ----------------------------------------------------------

  if (!canSupervisorTakeProfessor(supervisor, professorBundles)) {
    return false;
  }

  // ----------------------------------------------------------
  // Snapshot قبل التوزيع
  // ----------------------------------------------------------

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

  // ----------------------------------------------------------
  // نحاول نضيف كل Bundles
  // ----------------------------------------------------------

  for (const bundle of professorBundles) {
    const ok = assignBundleToSupervisor(
      supervisor,
      bundle,
      result,
      bundleAssignments,
    );

    if (!ok) {
      // ------------------------------------------------------
      // ROLLBACK
      // ------------------------------------------------------

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

  // ----------------------------------------------------------
  // كل الدكتور نجح
  // ----------------------------------------------------------

  professorAssignments.set(professorKey, id);

  supervisor.assignedProfessors.add(professorKey);

  console.log(
    `✅ Professor assigned completely: ${professorKey} -> Supervisor ${id}`,
  );

  return true;
}

// ============================================================
// Quota Map
// ============================================================

function buildQuotaMap(periodQuotas = []) {
  const quotaMap = new Map();

  for (const item of periodQuotas) {
    const supervisorId = Number(item.supervisor_id);

    const target = Number(item.target_periods);

    if (!Number.isFinite(supervisorId)) {
      continue;
    }

    if (!Number.isInteger(target) || target <= 0) {
      continue;
    }

    quotaMap.set(supervisorId, target);
  }

  return quotaMap;
}

// ============================================================
// Quota Score
// ============================================================

function getQuotaScore(
  supervisor,
  workload,
  target,
  globalMinimumReached = false,
) {
  const current = Number(supervisor.total || 0);

  const projected = current + workload;

  // ما عنده Target
  if (target === null || target === undefined) {
    return {
      score: 0,
      projected,
      distance: 0,
      deficit: 0,
      belowTarget: false,
    };
  }

  const deficit = Math.max(0, target - current);

  const projectedDeficit = Math.max(0, target - projected);

  const belowTarget = current < target;

  let score = 0;

  if (!globalMinimumReached) {
    if (belowTarget) {
      score += 1000000;

      score += deficit * 10000;

      score += (deficit - projectedDeficit) * 5000;
    } else {
      score -= 500000;
    }
  }

  return {
    score,

    projected,

    distance: Math.abs(target - projected),

    deficit,

    belowTarget,
  };
}

// ============================================================
// Generate Plan
// ============================================================

async function generatePlan(planId, variant = 1) {
  console.log(`🚀 Generating plan ${planId}, variant ${variant}`);

  // بنجيب بيانات الخطة
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
    periodQuotas = [],
  } = ctx;

  console.log("========================================");

  console.log("🔎 DEBUG PLAN DATA");

  console.log("========================================");

  console.log("📦 Groups from getPlanContext:", groups.length);

  const groupsByDate = {};

  for (const group of groups) {
    const day = dateISO(group.date);

    if (!groupsByDate[day]) {
      groupsByDate[day] = [];
    }

    groupsByDate[day].push(group);
  }

  console.log("📅 Groups by date:");

  for (const [day, dayGroups] of Object.entries(groupsByDate)) {
    console.log(day, "=>", dayGroups.length);
  }

  console.log("========================================");

  // ==========================================================
  // Validation
  // ==========================================================

  if (!groups.length) {
    await saveAssignments(planId, []);

    return {
      assigned: 0,
      total: 0,
      conflicts: [],
      supervisorsUsed: 0,
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

  // إذا ما في Duty Pool بنستخدم كل المشرفين
  if (!selectedSupervisorIds.length) {
    selectedSupervisorIds = supervisors
      .map((s) => Number(s.id))
      .filter(Number.isFinite);
  }

  // نشيل التكرار
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

      // مجموع الفترات
      total: 0,

      // آخر يوم اشتغل فيه
      lastDay: null,

      // Day -> عدد الفترات
      byDay: {},

      // Day -> Set periods
      byDayPeriods: {},

      // Day -> Set period ranks
      byDayPeriodRanks: {},

      // الأماكن المحجوزة
      occupiedSlots: new Set(),

      // الـ Session Groups
      assignedGroups: new Set(),

      // الدكاترة
      assignedProfessors: new Set(),
    };
  }

  // ==========================================================
  // Period Quotas
  // ==========================================================

  const quotaMap = buildQuotaMap(periodQuotas);

  console.log("🎯 Period quotas:", Object.fromEntries(quotaMap));

  function hasReachedAllQuotas(cand, selectedSupervisorIds, quotaMap) {
    // ما في quotas
    if (!quotaMap.size) {
      return true;
    }

    for (const supervisorId of selectedSupervisorIds) {
      const target = quotaMap.get(Number(supervisorId));

      if (target === null || target === undefined) {
        continue;
      }

      const supervisor = cand[Number(supervisorId)];

      if (!supervisor) {
        continue;
      }

      if (Number(supervisor.total || 0) < Number(target)) {
        return false;
      }
    }

    return true;
  }

  // ==========================================================
  // Affinities
  // ==========================================================

  const professorAffinity = new Map();

  const professorNameAffinity = new Map();

  console.log("========================================");

  console.log("🔗 LOADING PROFESSOR AFFINITIES");

  console.log("========================================");

  console.log("Affinity rows from DB:", aff.length);

  // ==========================================================
  // Result State
  // ==========================================================

  const result = [];

  const conflicts = [];

  const bundleAssignments = new Map();

  // ==========================================================
  // Load Affinities
  // ==========================================================

  for (const item of aff) {
    const supervisorId = Number(item.supervisor_id ?? item.supervisorId);

    if (!Number.isFinite(supervisorId)) {
      console.warn("⚠️ Affinity skipped: invalid supervisor ID", item);

      continue;
    }

    // --------------------------------------------------------
    // المشرف لازم يكون من Duty Pool
    // --------------------------------------------------------

    if (!cand[supervisorId]) {
      console.warn(
        `⚠️ Affinity supervisor ${supervisorId} is not in Duty Pool`,
      );

      conflicts.push({
        type: "AFFINITY_SUPERVISOR_NOT_SELECTED",

        professor: item.professor_name || item.name || "",

        professor_id: item.professor_id ?? null,

        supervisor_id: supervisorId,

        message: "The affinity supervisor is not in the selected Duty Pool.",
      });

      continue;
    }

    // --------------------------------------------------------
    // Professor ID
    // --------------------------------------------------------

    const professorId = Number(item.professor_id ?? item.professorId ?? null);

    if (Number.isFinite(professorId)) {
      professorAffinity.set(professorId, supervisorId);
    }

    // --------------------------------------------------------
    // Professor Name
    // --------------------------------------------------------

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

  console.log(`🔗 Professor ID affinities: ${professorAffinity.size}`);

  console.log(`🔗 Professor name affinities: ${professorNameAffinity.size}`);

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

        professor: group.professor_name || group.professor || "",

        professor_name: group.professor_name || group.professor || "",

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
  // Group Bundles By Professor
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

  // ==========================================================
  // Sort Bundles
  // ==========================================================

  for (const professor of professorGroups) {
    professor.bundles.sort((a, b) => {
      const dateCompare = String(a.date).localeCompare(String(b.date));

      if (dateCompare !== 0) {
        return dateCompare;
      }

      return sortPeriods(a.period, b.period);
    });
  }

  // ==========================================================
  // Sort Professors By Workload
  // ==========================================================

  professorGroups.sort((a, b) => {
    if (b.bundles.length !== a.bundles.length) {
      return b.bundles.length - a.bundles.length;
    }

    return String(a.professor).localeCompare(String(b.professor), "ar");
  });

  console.log(`👨‍🏫 Unique professors: ${professorGroups.length}`);

  // ==========================================================
  // Professor Assignments
  // ==========================================================

  const professorAssignments = new Map();

  // ==========================================================
  // Forced Assignments
  // ==========================================================

  const forcedProfessorAssignments = new Map();

  function registerForcedProfessor(group, supervisorId, source) {
    const professorKey = getProfessorKey(group);

    const id = Number(supervisorId);

    if (!Number.isFinite(id)) {
      return;
    }

    // المشرف مش مختار
    if (!cand[id]) {
      conflicts.push({
        type: "SUPERVISOR_NOT_SELECTED",

        professor: group.professor_name || group.professor || "",

        professor_id: group.professor_id ?? null,

        supervisor_id: id,

        source,

        message: "The required supervisor is not in the selected Duty Pool.",
      });

      return;
    }

    // أول Forced assignment
    if (!forcedProfessorAssignments.has(professorKey)) {
      forcedProfessorAssignments.set(professorKey, {
        supervisorId: id,
        source,
      });

      return;
    }

    const existing = forcedProfessorAssignments.get(professorKey);

    // نفس المشرف
    if (Number(existing.supervisorId) === id) {
      return;
    }

    // الدكتور مربوط بمشرفين مختلفين
    conflicts.push({
      type: "PROFESSOR_MULTIPLE_SUPERVISORS",

      professor: group.professor_name || group.professor || "",

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

    // --------------------------------------------------------
    // 1. نحاول Professor ID أولاً
    // --------------------------------------------------------

    if (
      professor.professor_id !== null &&
      professor.professor_id !== undefined
    ) {
      affinitySupervisor = professorAffinity.get(
        Number(professor.professor_id),
      );
    }

    // --------------------------------------------------------
    // 2. إذا ما لقيناه -> الاسم
    // --------------------------------------------------------

    if (affinitySupervisor === null || affinitySupervisor === undefined) {
      affinitySupervisor = professorNameAffinity.get(
        professor.professorNameKey,
      );
    }

    // --------------------------------------------------------
    // 3. وجدنا Affinity
    // --------------------------------------------------------

    if (affinitySupervisor !== null && affinitySupervisor !== undefined) {
      const supervisorId = Number(affinitySupervisor);

      console.log(
        `🎯 AFFINITY FOUND: Professor "${professor.professor}" -> Supervisor ${supervisorId}`,
      );

      // ------------------------------------------------------
      // تأكد أن المشرف موجود في الـ Duty Pool
      // ------------------------------------------------------

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

      // ------------------------------------------------------
      // هل عنده Lock / Preassignment؟
      // ------------------------------------------------------

      const forced = forcedProfessorAssignments.get(professor.key);

      // ------------------------------------------------------
      // تعارض
      // ------------------------------------------------------

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

    const ok = assignProfessorToSupervisor(
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
  // Shuffle
  // ==========================================================

  const shuffle = (arr) => seededShuffle([...arr], Number(variant) || 1);

  // ==========================================================
  // Remaining Professors
  // ==========================================================

  for (const professor of professorGroups) {
    // إذا موزع من قبل
    if (professorAssignments.has(professor.key)) {
      continue;
    }

    // بنشوف إذا كل الـ quotas وصلت
    const allQuotasReached = hasReachedAllQuotas(
      cand,
      selectedSupervisorIds,
      quotaMap,
    );

    const ranked = [];

    // بنفحص كل المشرفين
    for (const supervisorId of selectedSupervisorIds) {
      const supervisor = cand[supervisorId];

      if (!supervisor) {
        continue;
      }

      // الدكتور لازم يركب كامل عند المشرف
      if (!canSupervisorTakeProfessor(supervisor, professor.bundles)) {
        continue;
      }

      const workload = professor.bundles.length;

      const currentTotal = Number(supervisor.total || 0);

      const projectedTotal = currentTotal + workload;

      // Quota
      const quota = quotaMap.has(supervisorId)
        ? quotaMap.get(supervisorId)
        : null;

      const quotaInfo = getQuotaScore(
        supervisor,
        workload,
        quota,
        allQuotasReached,
      );

      // --------------------------------------------------------
      // الفترات المتتالية القديمة
      // --------------------------------------------------------

      const consecutiveScore = getProfessorConsecutiveScore(
        supervisor,
        professor.bundles,
      );

      // --------------------------------------------------------
      // ⭐ التعديل الرئيسي
      //
      // نحسب مدى اتصال الفترات بعد إضافة الدكتور
      // --------------------------------------------------------

      const periodContinuityScore = getPeriodContinuityScore(
        supervisor,
        professor.bundles,
      );

      // --------------------------------------------------------
      // ⭐ الأولوية النهائية للاستمرارية
      // --------------------------------------------------------

      const continuityPriority = getContinuityPriority(
        supervisor,
        professor.bundles,
      );

      // ضغط المشرف خلال الأيام
      const dailyLoad = getProfessorDailyLoad(supervisor, professor.bundles);

      // عدد الدكاترة عند المشرف
      const professorCount = supervisor.assignedProfessors?.size || 0;

      let score = 0;

      // ======================================================
      // MODE 2: Quota
      // ======================================================

      if (quota !== null) {
        const target = Number(quota);

        const currentDeficit = Math.max(0, target - currentTotal);

        const projectedDeficit = Math.max(0, target - projectedTotal);

        const reachesTarget = currentTotal < target && projectedTotal >= target;

        score =
          // ==================================================
          // ⭐ الاستمرارية أولاً
          // ==================================================

          periodContinuityScore * 1000000 +
          continuityPriority * 1000 +
          // المشرف تحت الهدف
          (currentTotal < target ? 100000000 : 0) +
          // نقربه من الهدف
          (currentDeficit - projectedDeficit) * 10000 +
          // وصل الهدف
          (reachesTarget ? 500000 : 0) -
          // عقوبة التجاوز
          Math.max(0, projectedTotal - target) * 1000 -
          // Fairness
          projectedTotal * 10 -
          // ضغط اليوم
          dailyLoad * 5;
      }

      // ======================================================
      // MODE 1: Everyone
      // ======================================================
      else {
        const noProfessorYet = professorCount === 0;

        score =
          // ==================================================
          // ⭐ الاستمرارية أولاً
          // ==================================================

          periodContinuityScore * 1000000 +
          continuityPriority * 1000 +
          // كل مشرف يأخذ دكتور أول
          (noProfessorYet ? 100000000 : 0) -
          // نوازن مجموع الفترات
          projectedTotal * 10000 -
          // ضغط اليوم
          dailyLoad * 2;
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

        quota,

        quotaScore: quotaInfo.score,

        quotaDistance: quotaInfo.distance,

        score,
      });
    }

    // ========================================================
    // No Supervisor
    // ========================================================

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
    for (const item of ranked) {
      if (item.quota !== null) {
        item.quotaOvershoot = Math.max(
          0,
          item.projectedTotal - Number(item.quota),
        );

        item.quotaGapAfterAssignment = Math.abs(
          Number(item.quota) - item.projectedTotal,
        );
      } else {
        item.quotaOvershoot = 0;
        item.quotaGapAfterAssignment = 0;
      }
    }

    // ========================================================
    // ⭐ SORT SUPERVISORS
    // ========================================================
    ranked.sort((a, b) => {
      // ======================================================
      // 1. QUOTA TARGET
      // ======================================================

      if (!allQuotasReached) {
        const aHasQuota = a.quota !== null && a.currentTotal < Number(a.quota);

        const bHasQuota = b.quota !== null && b.currentTotal < Number(b.quota);

        // المشرف الذي لم يصل Target بعد له أولوية
        if (aHasQuota !== bHasQuota) {
          return aHasQuota ? -1 : 1;
        }

        // إذا الاثنين تحت Target
        if (aHasQuota && bHasQuota) {
          const aDeficit = Math.max(0, Number(a.quota) - a.currentTotal);

          const bDeficit = Math.max(0, Number(b.quota) - b.currentTotal);

          // الأكبر deficit أولاً
          if (aDeficit !== bDeficit) {
            return bDeficit - aDeficit;
          }

          // الأقرب للهدف بعد إضافة الدكتور
          if (a.quotaDistance !== b.quotaDistance) {
            return a.quotaDistance - b.quotaDistance;
          }

          if (a.quotaGapAfterAssignment !== b.quotaGapAfterAssignment) {
            return a.quotaGapAfterAssignment - b.quotaGapAfterAssignment;
          }
        }
      }

      // ======================================================
      // 2. FAIRNESS
      // ======================================================

      if (a.projectedTotal !== b.projectedTotal) {
        return a.projectedTotal - b.projectedTotal;
      }

      // ======================================================
      // 3. CONTINUITY
      // ======================================================

      if (a.periodContinuityScore !== b.periodContinuityScore) {
        return b.periodContinuityScore - a.periodContinuityScore;
      }

      // ======================================================
      // 4. CONTINUITY PRIORITY
      // ======================================================

      if (a.continuityPriority !== b.continuityPriority) {
        return b.continuityPriority - a.continuityPriority;
      }

      // ======================================================
      // 5. CONSECUTIVE PERIODS
      // ======================================================

      if (a.consecutiveScore !== b.consecutiveScore) {
        return b.consecutiveScore - a.consecutiveScore;
      }

      // ======================================================
      // 6. DAILY LOAD
      // ======================================================

      if (a.dailyLoad !== b.dailyLoad) {
        return a.dailyLoad - b.dailyLoad;
      }

      // ======================================================
      // 7. NUMBER OF PROFESSORS
      // ======================================================

      if (a.professorCount !== b.professorCount) {
        return a.professorCount - b.professorCount;
      }

      // ======================================================
      // 8. RANDOMIZATION
      // ======================================================

      return (
        seededShuffle(`${variant}-${a.supervisorId}`) -
        seededShuffle(`${variant}-${b.supervisorId}`)
      );
    });

    // ========================================================
    // ⭐ DEBUG
    // ========================================================

    if (ranked[0].periodContinuityScore !== 0) {
      console.log("🏆 SELECTED SUPERVISOR:", {
        professor: professor.professor,

        supervisor: ranked[0].supervisorId,

        quota: ranked[0].quota,

        currentTotal: ranked[0].currentTotal,

        projectedTotal: ranked[0].projectedTotal,

        quotaDistance: ranked[0].quotaDistance,

        periodContinuityScore: ranked[0].periodContinuityScore,

        continuityPriority: ranked[0].continuityPriority,

        consecutiveScore: ranked[0].consecutiveScore,

        dailyLoad: ranked[0].dailyLoad,
      });
    }

    // ========================================================
    // أول واحد بعد الترتيب هو الأنسب
    // ========================================================

    const selected = ranked[0];

    // ========================================================
    // Assign Professor
    // ========================================================

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
  // Save Assignments
  // ==========================================================

  await saveAssignments(planId, result);

  // ==========================================================
  // Statistics
  // ==========================================================

  const assignedGroupIds = new Set(
    result.map((r) => Number(r.session_group_id)),
  );

  const assignedGroups = assignedGroupIds.size;

  const totalGroups = groups.length;

  const supervisorsUsed = Object.values(cand).filter((c) => c.total > 0);

  // Fairness
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

  for (const [professorKey, supervisorId] of professorSupervisorCheck) {
    const assigned = professorAssignments.get(professorKey);

    if (supervisorId === "MULTIPLE") {
      professorUniquenessViolations++;

      continue;
    }

    if (assigned !== undefined && Number(assigned) !== Number(supervisorId)) {
      professorUniquenessViolations++;
    }
  }

  // ==========================================================
  // Quota Statistics
  // ==========================================================

  const quotaStatistics = Object.values(cand).map((c) => {
    const supervisor = supervisors.find((s) => Number(s.id) === Number(c.id));

    const target = quotaMap.has(c.id) ? quotaMap.get(c.id) : null;

    return {
      "Supervisor ID": c.id,

      Supervisor: supervisor?.name || c.id,

      "Actual Periods": c.total,

      "Target Periods": target ?? "",

      Difference: target !== null ? c.total - target : "",
    };
  });

  // ==========================================================
  // Logs
  // ==========================================================

  console.log(
    `✅ Plan generated: ${assignedGroups}/${totalGroups} groups assigned`,
  );

  console.log(`📦 Bundles assigned: ${assignedBundles}/${totalBundles}`);

  console.log(`👨‍🏫 Unique professors: ${professorGroups.length}`);

  console.log(`👥 Supervisors used: ${supervisorsUsed.length}`);

  console.log(`⚠️ Conflicts: ${conflicts.length}`);

  console.log(`⚖️ Fairness difference: ${fairnessDifference}`);

  console.log(
    `🔒 Professor uniqueness violations: ${professorUniquenessViolations}`,
  );

  console.log("🎯 Quota statistics:", quotaStatistics);

  // ==========================================================
  // Export Excel
  // ==========================================================

  const exportDir = path.join(__dirname, "../exports");

  if (!fs.existsSync(exportDir)) {
    fs.mkdirSync(exportDir, {
      recursive: true,
    });
  }

  // ==========================================================
  // Distribution Sheet
  // ==========================================================

  const distributionRows = result.map((row) => {
    const supervisor = supervisors.find(
      (s) => Number(s.id) === Number(row.supervisor_id),
    );

    return {
      "Session Group ID": row.session_group_id,

      CRN: row.crn,

      Professor: row.professor,

      Date: row.date,

      Period: row.period,

      Supervisor: supervisor?.name || row.supervisor_id,
    };
  });

  // ==========================================================
  // Conflicts Sheet
  // ==========================================================

  const conflictsRows = conflicts.map((item) => ({
    Date: item.date || "",

    Period: item.period || "",

    Professor: item.professor || "",

    "Professor ID": item.professor_id || "",

    Type: item.type || "",

    "Supervisor ID": item.supervisor_id || "",

    Message: item.message || "",
  }));

  // ==========================================================
  // Statistics Sheet
  // ==========================================================

  const statisticsRows = Object.values(cand).map((c) => {
    const supervisor = supervisors.find((s) => Number(s.id) === Number(c.id));

    const target = quotaMap.has(c.id) ? quotaMap.get(c.id) : null;

    return {
      "Supervisor ID": c.id,

      Supervisor: supervisor?.name || c.id,

      "Total Periods": c.total,

      "Target Periods": target ?? "",

      "Difference From Target": target !== null ? c.total - target : "",

      "Used Days": Object.keys(c.byDay).length,
    };
  });

  // ==========================================================
  // Professor Assignment Sheet
  // ==========================================================

  const professorRows = professorGroups.map((professor) => {
    const supervisorId = professorAssignments.get(professor.key);

    const supervisor = supervisors.find(
      (s) => Number(s.id) === Number(supervisorId),
    );

    return {
      "Professor ID": professor.professor_id,

      Professor: professor.professor,

      "Total Periods": professor.bundles.length,

      Supervisor: supervisor?.name || supervisorId || "",
    };
  });

  // ==========================================================
  // Quota Sheet
  // ==========================================================

  const quotaRows = quotaStatistics;

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

  const quotasSheet = xlsx.utils.json_to_sheet(quotaRows);

  xlsx.utils.book_append_sheet(workbook, distributionSheet, "Distribution");

  xlsx.utils.book_append_sheet(workbook, conflictsSheet, "Conflicts");

  xlsx.utils.book_append_sheet(workbook, statisticsSheet, "Statistics");

  xlsx.utils.book_append_sheet(
    workbook,
    professorsSheet,
    "Professor Assignment",
  );

  xlsx.utils.book_append_sheet(workbook, quotasSheet, "Period Quotas");

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

    assigned: assignedGroups,

    total: totalGroups,

    assignedBundles,

    totalBundles,

    conflicts,

    conflictsCount: conflicts.length,

    supervisorsUsed: supervisorsUsed.length,

    fairnessDifference,

    idealFairnessRange: "0-1",

    professorUniquenessViolations,

    exportPath,

    statistics: statisticsRows,

    quotaStatistics,

    selectedSupervisors: selectedSupervisorIds,
  };
}

// ============================================================
// Daily Pools
// ============================================================

// بنخليها للتوافق مع باقي المشروع
// الخوارزمية الجديدة ما بتستخدمها كـ hard restriction
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
