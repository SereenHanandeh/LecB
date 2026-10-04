const pool = require("../models/db");

// ============================================================
// List
// ============================================================

async function listRooms({ onlyActive = false } = {}) {
  const result = await pool.query(
    `
    SELECT id, room_number, capacity, building, tag, categories, is_active, created_at
    FROM rooms
    ${onlyActive ? "WHERE is_active = TRUE" : ""}
    ORDER BY
      LPAD(
        COALESCE(
          NULLIF(regexp_replace(room_number, '\\D', '', 'g'), ''),
          '9999999999'
        ),
        10,
        '0'
      ),
      room_number
    `,
  );

  return result.rows;
}

// ============================================================
// Active room numbers only (used for validation elsewhere)
// ============================================================

async function getActiveRoomNumbers() {
  const result = await pool.query(
    `SELECT room_number FROM rooms WHERE is_active = TRUE`,
  );

  return new Set(result.rows.map((r) => String(r.room_number)));
}

const VALID_CATEGORIES = ["متطلبات", "دبلوم", "مدمج"];

function normalizeCategories(categories) {
  if (categories === undefined) return undefined;
  if (!Array.isArray(categories)) {
    const err = new Error("الفئات يجب أن تكون مصفوفة");
    err.status = 400;
    throw err;
  }
  const bad = categories.filter((c) => !VALID_CATEGORIES.includes(c));
  if (bad.length) {
    const err = new Error(`فئات غير صالحة: ${bad.join("، ")}`);
    err.status = 400;
    throw err;
  }
  return [...new Set(categories)];
}

// ============================================================
// Create
// ============================================================

async function createRoom({ roomNumber, tag, categories }) {
  const number = String(roomNumber ?? "").trim();

  if (!number) {
    const err = new Error("رقم/اسم القاعة مطلوب");
    err.status = 400;
    throw err;
  }

  try {
    const cats = normalizeCategories(categories) ?? [];
    const cleanTag = tag ? String(tag).trim() : null;

    const result = await pool.query(
      `
      INSERT INTO rooms (room_number, tag, categories)
      VALUES ($1, $2, $3)
      RETURNING *
      `,
      [number, cleanTag, cats],
    );
    return result.rows[0];
  } catch (error) {
    if (error.code === "23505") {
      const err = new Error(`القاعة "${number}" موجودة مسبقًا`);
      err.status = 409;
      throw err;
    }

    throw error;
  }
}

// ============================================================
// Update
// ============================================================

async function updateRoom(
  roomId,
  { roomNumber, capacity, building, isActive, tag, categories },
) {
  const id = Number(roomId);

  if (!Number.isInteger(id)) {
    const err = new Error("معرف القاعة غير صالح");
    err.status = 400;
    throw err;
  }

  const existing = await pool.query(`SELECT * FROM rooms WHERE id = $1`, [id]);

  if (!existing.rowCount) {
    const err = new Error("القاعة غير موجودة");
    err.status = 404;
    throw err;
  }

  const current = existing.rows[0];

  const nextRoomNumber =
    roomNumber !== undefined && roomNumber !== null && String(roomNumber).trim()
      ? String(roomNumber).trim()
      : current.room_number;

  const nextCapacity =
    capacity === undefined
      ? current.capacity
      : capacity === "" || capacity === null
        ? null
        : Number(capacity);

  const nextBuilding =
    building === undefined
      ? current.building
      : building
        ? String(building).trim()
        : null;

  const nextIsActive =
    isActive === undefined ? current.is_active : Boolean(isActive);

  const nextTag =
    tag === undefined ? current.tag : tag ? String(tag).trim() : null;

  const nextCategories =
    categories === undefined
      ? current.categories
      : (normalizeCategories(categories) ?? current.categories);

  try {
    const result = await pool.query(
      `
      UPDATE rooms
      SET room_number = $2,
          capacity = $3,
          building = $4,
          is_active = $5,
          tag = $6,
          categories = $7
      WHERE id = $1
      RETURNING *
      `,
      [
        id,
        nextRoomNumber,
        nextCapacity,
        nextBuilding,
        nextIsActive,
        nextTag,
        nextCategories,
      ],
    );

    return result.rows[0];
  } catch (error) {
    if (error.code === "23505") {
      const err = new Error(`القاعة "${nextRoomNumber}" موجودة مسبقًا`);
      err.status = 409;
      throw err;
    }

    throw error;
  }
}

// ============================================================
// Delete
// ============================================================

async function deleteRoom(roomId) {
  const id = Number(roomId);

  if (!Number.isInteger(id)) {
    const err = new Error("معرف القاعة غير صالح");
    err.status = 400;
    throw err;
  }

  const result = await pool.query(
    `DELETE FROM rooms WHERE id = $1 RETURNING id`,
    [id],
  );

  if (!result.rowCount) {
    const err = new Error("القاعة غير موجودة");
    err.status = 404;
    throw err;
  }

  return result.rows[0];
}

module.exports = {
  listRooms,
  getActiveRoomNumbers,
  createRoom,
  updateRoom,
  deleteRoom,
};