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
    const { roomNumber, capacity, building } = req.body;

    const room = await createRoom({ roomNumber, capacity, building });

    return res.status(201).json({ success: true, data: room });
  } catch (err) {
    return handleError(res, err, "Error creating room");
  }
}

async function editRoom(req, res) {
  try {
    const { roomId } = req.params;

    const { roomNumber, capacity, building, isActive } = req.body;

    const room = await updateRoom(roomId, {
      roomNumber,
      capacity,
      building,
      isActive,
    });

    return res.json({ success: true, data: room });
  } catch (err) {
    return handleError(res, err, "Error updating room");
  }
}

async function removeRoom(req, res) {
  try {
    const { roomId } = req.params;

    const deleted = await deleteRoom(roomId);

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