const express = require("express");
const C = require('../controllers/supervisors.js');
const supervisorRouter = express.Router();

supervisorRouter.get('/', C.list);
supervisorRouter.post('/', C.create);
supervisorRouter.patch('/:id', C.update);

module.exports = supervisorRouter;
