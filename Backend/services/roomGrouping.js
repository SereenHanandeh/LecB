// يعيد توزيع القاعات بعد توزيع المشرفين:
// دكاترة نفس المشرف بقاعة وحدة إذا ما في تداخل بالفترات، وإذا في تداخل بقاعة ثانية.
// القاعات المثبّتة (pinned) ما بتتغير، وكل قاعة تخدم دكتور واحد بالفترة الواحدة.

function regroupRoomsBySupervisor({
  groups,
  result,
  rooms,
  pool,
  dateISO,
  normalizePeriod,
}) {
  const roomPool = [...new Set((pool || []).map(String))]; // مرتبة بالأولوية من الواجهة
  const rank = (room) => {
    const i = roomPool.indexOf(room);
    return i === -1 ? 9999 : i;
  };
  const byRank = (a, b) => rank(a) - rank(b);

  const conflicts = [];

  // مشرف كل محاضرة
  const supByGroup = new Map(
    result.map((r) => [Number(r.session_group_id), Number(r.supervisor_id)]),
  );

  // الفترات (أستاذ + تاريخ + فترة)
  const slots = new Map();

  for (const g of groups) {
    const pid = Number(g.professor_id);
    if (!Number.isFinite(pid)) continue;

    const date = dateISO(g.date);
    const periodNorm = normalizePeriod(g.period_label);
    if (!date || !periodNorm) continue;

    const key = `${pid}|${date}|${periodNorm}`;

    if (!slots.has(key)) {
      slots.set(key, {
        key,
        pid,
        date,
        periodNorm,
        periodRaw: g.period_label,
        name: g.professor_name ?? "",
        supervisorId: supByGroup.get(Number(g.id)) ?? null,
      });
    }
  }

  // القاعات الأصلية والمثبّتة
  const originalRoom = new Map();
  const pinnedKeys = new Set();
  const assigned = new Map();
  const occupied = new Map(); // date|period|room -> professorId

  const occKey = (slot, room) => `${slot.date}|${slot.periodNorm}|${room}`;

  for (const item of rooms) {
    const pid = Number(item.professor_id);
    const date = dateISO(item.date ?? "");
    const period = normalizePeriod(item.period_label ?? "");
    const room = String(item.room_number ?? "").trim();

    if (!Number.isFinite(pid) || !date || !period || !room) continue;

    const key = `${pid}|${date}|${period}`;
    originalRoom.set(key, room);

    if (item.pinned && slots.has(key)) {
      pinnedKeys.add(key);
      assigned.set(key, room);
      occupied.set(`${date}|${period}|${room}`, pid);
    }
  }

  // تجميع حسب المشرف ثم الأستاذ
  const bySupervisor = new Map();

  for (const slot of slots.values()) {
    if (pinnedKeys.has(slot.key)) continue;

    const supKey = slot.supervisorId ?? "none";
    if (!bySupervisor.has(supKey)) bySupervisor.set(supKey, new Map());

    const byProf = bySupervisor.get(supKey);
    if (!byProf.has(slot.pid)) byProf.set(slot.pid, []);
    byProf.get(slot.pid).push(slot);
  }

  const supervisorOrder = [...bySupervisor.values()]
    .map((byProf) => ({
      byProf,
      total: [...byProf.values()].reduce((sum, list) => sum + list.length, 0),
    }))
    .sort((a, b) => b.total - a.total);

  const slotFree = (slot, room) => !occupied.has(occKey(slot, room));

  const take = (slot, room) => {
    assigned.set(slot.key, room);
    occupied.set(occKey(slot, room), slot.pid);
  };

  for (const { byProf } of supervisorOrder) {
    let supRooms = []; // قاعات هذا المشرف

    const professorsList = [...byProf.entries()].sort(
      (a, b) => b[1].length - a[1].length,
    );

    for (const [, profSlots] of professorsList) {
      const allFree = (room) => profSlots.every((s) => slotFree(s, room));

      // 1) قاعة المشرف الحالية إذا ما في تداخل، 2) وإلا قاعة جديدة
      const room = supRooms.find(allFree) ?? roomPool.find(allFree);

      if (room) {
        profSlots.forEach((s) => take(s, room));

        if (!supRooms.includes(room)) {
          supRooms = [...supRooms, room].sort(byRank);
        }

        continue;
      }

      // احتياط: قاعة مختلفة لكل فترة
          // احتياط: قاعة مختلفة لكل فترة
      for (const s of profSlots) {
        const r =
          supRooms.find((x) => slotFree(s, x)) ??
          roomPool.find((x) => slotFree(s, x));

        if (r) {
          take(s, r);

          if (!supRooms.includes(r)) {
            supRooms = [...supRooms, r].sort(byRank);
          }

          continue;
        }

        // القاعة الأصلية فقط إذا كانت فاضية (حتى لا تتكرر القاعة بنفس الفترة)
        const original = originalRoom.get(s.key);
        if (original && slotFree(s, original)) {
          take(s, original);
          continue;
        }

        conflicts.push({
          type: "ROOM_NOT_AVAILABLE",
          professor: s.name,
          professor_id: s.pid,
          date: s.date,
          period: s.periodNorm,
          message: "لا توجد قاعة فاضية بهذه الفترة ضمن القاعات المتاحة.",
        });
      }
    }
  }

  const rows = [];
  let changed = 0;

  for (const slot of slots.values()) {
    const room = assigned.get(slot.key);
    if (!room) continue;

    if (originalRoom.get(slot.key) !== room) changed++;

    rows.push({
      professor_id: slot.pid,
      name: slot.name,
      room_number: room,
      date: slot.date,
      period_label: slot.periodRaw,
      pinned: pinnedKeys.has(slot.key),
    });
  }

  return { rows, conflicts, changed };
}

module.exports = { regroupRoomsBySupervisor };