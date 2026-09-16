const pool = require('../models/db');

// Get all supervisors
async function list(req, res) {
  try {
    const result = await pool.query('SELECT * FROM supervisors ORDER BY id');
    res.json({
      message: 'Supervisors fetched successfully',
      supervisors: result.rows
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server Error' });
  }
}

// Create a new supervisor
async function create(req, res) {
  try {
    const { name, active = true } = req.body;
    const result = await pool.query(
      'INSERT INTO supervisors(name, active) VALUES ($1,$2) RETURNING *',
      [name, active]
    );
    res.json({ message: 'Supervisor created successfully', supervisor: result.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server Error' });
  }
}

// Update an existing supervisor
async function update(req, res) {
  try {
    const { id } = req.params;
    const { name, active } = req.body;
    const result = await pool.query(
      'UPDATE supervisors SET name=COALESCE($1,name), active=COALESCE($2,active) WHERE id=$3 RETURNING *',
      [name, active, id]
    );
    res.json({ message: 'Supervisor updated successfully', supervisor: result.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server Error' });
  }
}

module.exports = {
  list,
  create,
  update
};
