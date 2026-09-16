// =====================================================
// Distribution Utilities
// =====================================================
//
// القواعد الأساسية:
//
// 1. العدالة حسب عدد الفترات.
// 2. الفرق المثالي بين المشرفين 0 أو 1.
// 3. نفس المشرف لا يأخذ نفس الفترة مرتين في نفس اليوم.
// 4. الدكتور الواحد يتم التعامل معه كوحدة واحدة.
// 5. Professor Affinity يعتمد على professor_id.
// 6. نفضل الفترات المتتالية.
// 7. نحاول تكوين بلوكات 3 فترات أو أكثر.
// =====================================================


// =====================================================
// Normalize Period
// =====================================================

function normalizePeriod(period) {
  if (
    period === null ||
    period === undefined
  ) {
    return "";
  }

  return String(period)
    .trim()
    .replace(/\s+/g, " ");
}


// =====================================================
// Period Rank
// =====================================================

function getPeriodRank(period) {
  const value =
    normalizePeriod(period);

  if (!value) {
    return null;
  }

  const arabicPeriods = {
    "الأولى": 1,
    "الاولى": 1,
    "الأول": 1,
    "الاول": 1,

    "الثانية": 2,
    "الثانيه": 2,
    "الثاني": 2,

    "الثالثة": 3,
    "الثالثه": 3,
    "الثالث": 3,

    "الرابعة": 4,
    "الرابعه": 4,
    "الرابع": 4,

    "الخامسة": 5,
    "الخامسه": 5,
    "الخامس": 5,

    "السادسة": 6,
    "السادسه": 6,
    "السادس": 6,

    "السابعة": 7,
    "السابعه": 7,
    "السابع": 7,

    "الثامنة": 8,
    "الثامنه": 8,
    "الثامن": 8,

    "التاسعة": 9,
    "التاسعه": 9,
    "التاسع": 9,

    "العاشرة": 10,
    "العاشره": 10,
    "العاشر": 10,
  };

  if (
    Object.prototype.hasOwnProperty.call(
      arabicPeriods,
      value
    )
  ) {
    return arabicPeriods[value];
  }

  const numeric =
    Number(value);

  if (
    Number.isInteger(numeric) &&
    numeric > 0
  ) {
    return numeric;
  }

  return null;
}


// =====================================================
// Build Daily Pools
// =====================================================

function buildDailyPools(
  days,
  supIds,
  cap = 10
) {
  const pools = {};

  const ids =
    Array.from(
      new Set(
        supIds
          .map((id) =>
            Number(id)
          )
          .filter((id) =>
            Number.isInteger(id)
          )
      )
    );

  if (
    !ids.length ||
    !days.length
  ) {
    return pools;
  }

  const dailyCount =
    Math.min(
      cap,
      ids.length
    );

  let previousDay =
    new Set();

  let order =
    [...ids];

  for (
    let dayIndex = 0;
    dayIndex < days.length;
    dayIndex++
  ) {
    const day =
      days[dayIndex];

    let candidates =
      order.filter(
        (id) =>
          !previousDay.has(id)
      );

    if (
      candidates.length <
      dailyCount
    ) {
      const fallback =
        order.filter(
          (id) =>
            !candidates.includes(
              id
            )
        );

      candidates = [
        ...candidates,
        ...fallback,
      ];
    }

    const today =
      candidates.slice(
        0,
        dailyCount
      );

    pools[day] =
      today;

    previousDay =
      new Set(today);

    if (
      order.length >
      dailyCount
    ) {
      order = [
        ...order.slice(
          dailyCount
        ),
        ...order.slice(
          0,
          dailyCount
        ),
      ];
    }
  }

  return pools;
}


// =====================================================
// Consecutive Score
// =====================================================

function getConsecutiveScore(
  candidate,
  day,
  period
) {
  const rank =
    getPeriodRank(period);

  if (
    rank === null
  ) {
    return 0;
  }

  const periods =
    candidate
      ?.byDayPeriodRanks
      ?.[
        day
      ] ||
    new Set();

  let score = 0;

  if (
    periods.has(
      rank - 1
    )
  ) {
    score += 100;
  }

  if (
    periods.has(
      rank + 1
    )
  ) {
    score += 40;
  }

  if (
    periods.has(
      rank - 1
    ) &&
    periods.has(
      rank - 2
    )
  ) {
    score += 80;
  }

  if (
    periods.has(
      rank - 1
    ) &&
    periods.has(
      rank - 2
    ) &&
    periods.has(
      rank - 3
    )
  ) {
    score += 60;
  }

  return score;
}


// =====================================================
// Current Consecutive Block
// =====================================================

function getCurrentBlockLength(
  candidate,
  day
) {
  const periods =
    candidate
      ?.byDayPeriodRanks
      ?.[
        day
      ];

  if (
    !periods ||
    periods.size === 0
  ) {
    return 0;
  }

  const values =
    Array.from(periods)
      .filter((x) =>
        Number.isInteger(x)
      )
      .sort(
        (a, b) =>
          a - b
      );

  if (!values.length) {
    return 0;
  }

  let best = 1;
  let current = 1;

  for (
    let i = 1;
    i < values.length;
    i++
  ) {
    if (
      values[i] ===
      values[i - 1] + 1
    ) {
      current++;

      if (
        current > best
      ) {
        best = current;
      }
    } else {
      current = 1;
    }
  }

  return best;
}


// =====================================================
// Pick Best Candidate
// =====================================================
//
// This function is kept for compatibility with other parts
// of the project.
//
// The main generatePlan now assigns PROFESSORS as a whole,
// so this function is no longer responsible for professor
// uniqueness.
// =====================================================

function pickBestCandidate({
  day,
  period,
  sg,
  dailyPool,
  cand,
  affMap,
}) {
  if (
    !dailyPool ||
    dailyPool.length === 0
  ) {
    return null;
  }

  const normalizedPeriod =
    normalizePeriod(period);

  const periodRank =
    getPeriodRank(
      normalizedPeriod
    );

  const eligible =
    dailyPool.filter(
      (id) => {
        const c =
          cand[id];

        if (!c) {
          return false;
        }

        const assignedGroups =
          c.assignedGroups ||
          new Set();

        if (
          sg &&
          sg.id !== undefined &&
          assignedGroups.has(
            Number(sg.id)
          )
        ) {
          return false;
        }

        const periods =
          c.byDayPeriods?.[
            day
          ] ||
          new Set();

        if (
          normalizedPeriod &&
          periods.has(
            normalizedPeriod
          )
        ) {
          return false;
        }

        if (
          periodRank !== null
        ) {
          const ranks =
            c.byDayPeriodRanks?.[
              day
            ] ||
            new Set();

          if (
            ranks.has(
              periodRank
            )
          ) {
            return false;
          }
        }

        return true;
      }
    );

  if (
    !eligible.length
  ) {
    return null;
  }

  const ranked =
    eligible.map(
      (id) => {
        const c =
          cand[id];

        let professorScore = 0;

        // ---------------------------------------------------
        // Map is supported.
        // ---------------------------------------------------

        let affinity = null;

        if (
          affMap instanceof Map
        ) {
          affinity =
            affMap.get(id);
        }

        if (
          affinity &&
          sg?.professor_id !==
            null &&
          sg?.professor_id !==
            undefined
        ) {
          if (
            affinity.professorIds?.has?.(
              Number(
                sg.professor_id
              )
            )
          ) {
            professorScore =
              100000;
          }
        }

        const total =
          Number(
            c.total || 0
          );

        const dayCount =
          Number(
            c.byDay?.[
              day
            ] || 0
          );

        const fairnessScore =
          -total * 1000;

        const consecutiveScore =
          getConsecutiveScore(
            c,
            day,
            normalizedPeriod
          );

        const blockLength =
          getCurrentBlockLength(
            c,
            day
          );

        const blockScore =
          blockLength >= 3
            ? 300
            : blockLength >= 2
            ? 180
            : 0;

        const dailyLoadScore =
          -dayCount * 30;

        const score =
          professorScore +
          fairnessScore +
          consecutiveScore +
          blockScore +
          dailyLoadScore;

        return {
          candidate: c,
          id,
          score,
          total,
          dayCount,
          professorScore,
          consecutiveScore,
          blockLength,
        };
      }
    );

  ranked.sort(
    (a, b) => {
      if (
        a.professorScore !==
        b.professorScore
      ) {
        return (
          b.professorScore -
          a.professorScore
        );
      }

      if (
        a.total !==
        b.total
      ) {
        return (
          a.total -
          b.total
        );
      }

      if (
        a.consecutiveScore !==
        b.consecutiveScore
      ) {
        return (
          b.consecutiveScore -
          a.consecutiveScore
        );
      }

      if (
        a.blockLength !==
        b.blockLength
      ) {
        return (
          b.blockLength -
          a.blockLength
        );
      }

      if (
        a.dayCount !==
        b.dayCount
      ) {
        return (
          a.dayCount -
          b.dayCount
        );
      }

      return (
        Number(a.id) -
        Number(b.id)
      );
    }
  );

  return (
    ranked[0]?.candidate ||
    null
  );
}


// =====================================================
// Assign Once
// =====================================================

function assignOnce(
  result,
  cand,
  supId,
  sg
) {
  const supervisorId =
    Number(supId);

  const c =
    cand[supervisorId];

  if (
    !c ||
    !sg
  ) {
    return false;
  }

  const groupId =
    Number(sg.id);

  if (
    !Number.isInteger(
      groupId
    )
  ) {
    return false;
  }

  c.assignedGroups =
    c.assignedGroups ||
    new Set();

  if (
    c.assignedGroups.has(
      groupId
    )
  ) {
    return false;
  }

  let day;

  if (
    sg.date instanceof Date
  ) {
    day =
      sg.date
        .toISOString()
        .slice(0, 10);
  } else {
    const text =
      String(
        sg.date || ""
      ).trim();

    if (
      /^\d{4}-\d{2}-\d{2}$/.test(
        text
      )
    ) {
      day = text;
    } else {
      const parsed =
        new Date(text);

      if (
        Number.isNaN(
          parsed.getTime()
        )
      ) {
        return false;
      }

      day =
        parsed
          .toISOString()
          .slice(0, 10);
    }
  }

  c.byDay =
    c.byDay || {};

  c.byDayPeriods =
    c.byDayPeriods || {};

  c.byDayPeriodRanks =
    c.byDayPeriodRanks || {};

  c.byDay[day] =
    Number(
      c.byDay[day] || 0
    );

  c.byDayPeriods[day] =
    c.byDayPeriods[day] ||
    new Set();

  c.byDayPeriodRanks[day] =
    c.byDayPeriodRanks[day] ||
    new Set();

  const period =
    normalizePeriod(
      sg.period_label
    );

  const periodRank =
    getPeriodRank(period);

  if (
    period &&
    c.byDayPeriods[day].has(
      period
    )
  ) {
    return false;
  }

  if (
    periodRank !== null &&
    c.byDayPeriodRanks[
      day
    ].has(periodRank)
  ) {
    return false;
  }

  c.assignedGroups.add(
    groupId
  );

  if (period) {
    c.byDayPeriods[day].add(
      period
    );
  }

  if (
    periodRank !== null
  ) {
    c.byDayPeriodRanks[
      day
    ].add(periodRank);
  }

  c.byDay[day] =
    Number(
      c.byDay[day] || 0
    ) + 1;

  c.total =
    Number(
      c.total || 0
    ) + 1;

  c.lastDay = day;

  result.push({
    session_group_id:
      groupId,

    supervisor_id:
      supervisorId,
  });

  return true;
}


// =====================================================
// Balance Fairness
// =====================================================

function balanceFairness(
  cand,
  result,
  diff = 1,
  supervisorIds = null
) {
  let arr =
    Object.values(cand);

  if (
    Array.isArray(
      supervisorIds
    ) &&
    supervisorIds.length
  ) {
    const ids =
      new Set(
        supervisorIds.map(
          Number
        )
      );

    arr =
      arr.filter(
        (c) =>
          ids.has(
            Number(c.id)
          )
      );
  }

  if (!arr.length) {
    return {
      balanced: true,
      difference: 0,
      min: 0,
      max: 0,
      target: 0,
      counts: [],
    };
  }

  const counts =
    arr.map(
      (c) => ({
        id:
          Number(c.id),

        total:
          Number(
            c.total || 0
          ),
      })
    );

  const values =
    counts.map(
      (x) =>
        x.total
    );

  const min =
    Math.min(
      ...values
    );

  const max =
    Math.max(
      ...values
    );

  const difference =
    max - min;

  const total =
    values.reduce(
      (sum, value) =>
        sum + value,
      0
    );

  const target =
    arr.length
      ? total / arr.length
      : 0;

  return {
    balanced:
      difference <= diff,

    difference,

    min,

    max,

    target,

    counts,
  };
}


// =====================================================
// Group By Polyfill
// =====================================================

function groupByPoly(
  arr,
  fn
) {
  const m = {};

  for (const e of arr) {
    const k = fn(e);

    if (!m[k]) {
      m[k] = [];
    }

    m[k].push(e);
  }

  return m;
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

  getPeriodRank,

  normalizePeriod,
};