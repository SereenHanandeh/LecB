const seededShuffle = require("./rng.js");

// =====================================================
// Group By
// =====================================================

function groupByPoly(array, keyFn) {
  return array.reduce((acc, item) => {
    const key = keyFn(item);

    if (!acc[key]) {
      acc[key] = [];
    }

    acc[key].push(item);

    return acc;
  }, {});
}

// =====================================================
// Build Daily Pools
// =====================================================

function buildDailyPools(
  days,
  supervisorIds,
  dailyCount = 10,
  variant = 1
) {
  const pools = {};

  const ids = Array.from(
    new Set(
      supervisorIds
        .map(Number)
        .filter((id) => Number.isInteger(id))
    )
  );

  if (!ids.length) {
    return pools;
  }

  if (ids.length < dailyCount) {
    throw new Error(
      `Need at least ${dailyCount} supervisors. Available: ${ids.length}`
    );
  }

  let order = seededShuffle(
    [...ids],
    Number(variant) || 1
  );

  let previousDay = new Set();

  days.forEach((day, dayIndex) => {
    // -----------------------------------------------
    // المشرفون الذين لم يعملوا أمس
    // -----------------------------------------------

    let candidates = order.filter(
      (id) => !previousDay.has(id)
    );


    if (candidates.length < dailyCount) {
      const remaining = order.filter(
        (id) => !candidates.includes(id)
      );

      candidates = [
        ...candidates,
        ...remaining,
      ];
    }

    // -----------------------------------------------
    // Shuffle
    // -----------------------------------------------

    candidates = seededShuffle(
      candidates,
      Number(variant) + dayIndex
    );

    // -----------------------------------------------
    // اختيار مشرفي اليوم
    // -----------------------------------------------

    const today = candidates.slice(
      0,
      dailyCount
    );

    pools[day] = today;

    previousDay = new Set(today);

    // -----------------------------------------------
    // Rotate
    // -----------------------------------------------

    order = seededShuffle(
      [
        ...order.slice(dailyCount),
        ...order.slice(0, dailyCount),
      ],
      Number(variant) + dayIndex + 100
    );
  });

  return pools;
}

// =====================================================
// Candidate Score
// =====================================================

function pickBestCandidate({
  day,
  period,
  sg,
  dailyPool = [],
  cand,
  affMap,
}) {
  if (!dailyPool.length) {
    return null;
  }

  const candidates = [];

  for (const supervisorId of dailyPool) {
    const id = Number(supervisorId);

    const state = cand[id];

    if (!state) {
      continue;
    }

    // -----------------------------------------------
    // منع المشرف من أخذ نفس الفترة مرتين
    // -----------------------------------------------

    const dayPeriods =
      state.byDayPeriods[day] || new Set();

    if (dayPeriods.has(period)) {
      continue;
    }

    // -----------------------------------------------
    // منع العمل يومين متتاليين
    // -----------------------------------------------

    if (
      state.lastDay &&
      isNextDay(state.lastDay, day)
    ) {
      continue;
    }

    // -----------------------------------------------
    // Affinity
    // -----------------------------------------------

    const affinity =
      affMap.get(id);

    let affinityScore = 0;

    if (affinity) {
      if (
        sg.crn &&
        affinity.crn.has(
          String(sg.crn)
        )
      ) {
        affinityScore += 100;
      }

      if (
        sg.professor_id &&
        affinity.prof.has(
          Number(sg.professor_id)
        )
      ) {
        affinityScore += 80;
      }
    }

    // -----------------------------------------------
    // Current day load
    // -----------------------------------------------

    const dailyLoad =
      state.byDay[day] || 0;

    // -----------------------------------------------
    // Total load
    // -----------------------------------------------

    const totalLoad =
      state.total || 0;

    // -----------------------------------------------
    // Score
    //
    // كلما قل score كان أفضل.
    // affinity تقلل score.
    // -----------------------------------------------

    const score =
      totalLoad * 10 +
      dailyLoad * 30 -
      affinityScore;

    candidates.push({
      id,
      score,
      totalLoad,
      dailyLoad,
      affinityScore,
    });
  }

  if (!candidates.length) {
    return null;
  }

  // -----------------------------------------------
  // أفضل مرشح
  // -----------------------------------------------

  candidates.sort((a, b) => {
    if (a.score !== b.score) {
      return a.score - b.score;
    }

    if (a.totalLoad !== b.totalLoad) {
      return a.totalLoad - b.totalLoad;
    }

    if (a.dailyLoad !== b.dailyLoad) {
      return a.dailyLoad - b.dailyLoad;
    }

    return a.id - b.id;
  });

  return candidates[0];
}

// =====================================================
// Assign Once
// =====================================================

function assignOnce(
  result,
  cand,
  supervisorId,
  sg
) {
  const id = Number(supervisorId);

  if (!cand[id]) {
    return false;
  }

  const groupId = Number(sg.id);

  const day = dateISO(sg.date);

  const period = sg.period_label;

  if (!day) {
    return false;
  }

  // -----------------------------------------------
  // منع التكرار لنفس Session Group
  // -----------------------------------------------

  const alreadyAssigned =
    result.some(
      (item) =>
        Number(item.session_group_id) ===
          groupId &&
        Number(item.supervisor_id) ===
          id
    );

  if (alreadyAssigned) {
    return false;
  }

  // -----------------------------------------------
  // منع نفس المشرف من نفس الفترة
  // -----------------------------------------------

  const dayPeriods =
    cand[id].byDayPeriods[day] ||
    new Set();

  if (dayPeriods.has(period)) {
    return false;
  }

  // -----------------------------------------------
  // تحديث state
  // -----------------------------------------------

  cand[id].total += 1;

  cand[id].byDay[day] =
    (cand[id].byDay[day] || 0) + 1;

  if (!cand[id].byDayPeriods[day]) {
    cand[id].byDayPeriods[day] =
      new Set();
  }

  cand[id].byDayPeriods[day].add(
    period
  );

  cand[id].lastDay = day;

  // -----------------------------------------------
  // إضافة النتيجة
  // -----------------------------------------------

  result.push({
    session_group_id: groupId,
    supervisor_id: id,
  });

  return true;
}



function balanceFairness(
  cand,
  result,
  allowedDifference = 2
) {
  const supervisors = Object.values(cand);

  if (supervisors.length < 2) {
    return result;
  }

  const loads = supervisors.map(
    (s) => s.total || 0
  );

  const maxLoad = Math.max(...loads);
  const minLoad = Math.min(...loads);

  const difference =
    maxLoad - minLoad;

  if (
    difference <=
    allowedDifference
  ) {
    return result;
  }

  console.warn(
    `⚠️ Fairness difference is ${difference}. Target <= ${allowedDifference}`
  );

  return result;
}

// =====================================================
// Date Helper
// =====================================================

function dateISO(value) {
  if (!value) {
    return null;
  }

  const date = new Date(value);

  if (Number.isNaN(date.getTime())) {
    return null;
  }

  return date
    .toISOString()
    .slice(0, 10);
}

// =====================================================
// Next Day
// =====================================================

function isNextDay(
  previousDate,
  currentDate
) {
  const previous = new Date(
    `${previousDate}T00:00:00`
  );

  const current = new Date(
    `${currentDate}T00:00:00`
  );

  const diff =
    current.getTime() -
    previous.getTime();

  return (
    diff ===
    24 * 60 * 60 * 1000
  );
}

// =====================================================
// Exports
// =====================================================

module.exports = {
  buildDailyPools,
  pickBestCandidate,
  assignOnce,
  balanceFairness,
  groupByPoly,
};
