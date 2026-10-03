const {
  listRooms,
  createRoom,
  updateRoom,
  deleteRoom,
} = require("../services/room.js");

function handleError(res, err, message = "Internal server error") {
  console.error(`❌ ${message}:`, err);

  return res.status(err.status || 500).json({
    success: false,
    error: err.message || message,
  });
}

async function getRooms(req, res) {
  try {
    const onlyActive = req.query.active === "true";

    const rooms = await listRooms({ onlyActive });

    return res.json({ success: true, data: rooms });
  } catch (err) {
    return handleError(res, err, "Error fetching rooms");
  }
}

async function addRoom(req, res) {
  try {
    const { roomNumber, capacity, building, tag, categories } = req.body;

    if (!roomNumber || !String(roomNumber).trim()) {
      return res.status(400).json({
        success: false,
        error: "رقم/اسم القاعة مطلوب",
      });
    }

    const normalizedCategories = Array.isArray(categories)
      ? categories.filter(Boolean)
      : [];

    const room = await createRoom({
      roomNumber: String(roomNumber).trim(),
      capacity: capacity === "" || capacity == null ? null : Number(capacity),
      building: building ? String(building).trim() : null,
      tag: tag ? String(tag).trim() : null,
      categories: normalizedCategories,
    });

    return res.status(201).json({ success: true, data: room });
  } catch (err) {
    return handleError(res, err, "Error creating room");
  }
}

async function editRoom(req, res) {
  try {
    const { roomId } = req.params;

    const { roomNumber, capacity, building, isActive, tag, categories } =
      req.body;

    const payload = {};

    if (roomNumber !== undefined) payload.roomNumber = String(roomNumber).trim();
    if (capacity !== undefined)
      payload.capacity = capacity === "" || capacity == null ? null : Number(capacity);
    if (building !== undefined)
      payload.building = building ? String(building).trim() : null;
    if (isActive !== undefined) payload.isActive = Boolean(isActive);
    if (tag !== undefined) payload.tag = tag ? String(tag).trim() : null;
    if (categories !== undefined)
      payload.categories = Array.isArray(categories)
        ? categories.filter(Boolean)
        : [];

    const room = await updateRoom(roomId, payload);

    if (!room) {
      return res.status(404).json({
        success: false,
        error: "القاعة غير موجودة.",
      });
    }

    return res.json({ success: true, data: room });
  } catch (err) {
    return handleError(res, err, "Error updating room");
  }
}

async function removeRoom(req, res) {
  try {
    const { roomId } = req.params;

    const deleted = await deleteRoom(roomId);

    if (!deleted) {
      return res.status(404).json({
        success: false,
        error: "القاعة غير موجودة.",
      });
    }

    return res.json({ success: true, data: deleted });
  } catch (err) {
    return handleError(res, err, "Error deleting room");
  }
}

module.exports = {
  getRooms,
  addRoom,
  editRoom,
  removeRoom,
};