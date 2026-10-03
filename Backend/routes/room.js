const express = require("express");

const R = require("../controllers/room.js");

const roomsRouter = express.Router();

roomsRouter.get("/",R.getRooms);
roomsRouter.post("/", R.addRoom);
roomsRouter.put("/:roomId", R.editRoom);
roomsRouter.delete("/:roomId",R.removeRoom);

module.exports = roomsRouter;