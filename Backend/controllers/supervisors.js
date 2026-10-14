const pool = require('../models/db');

// Get all supervisors (?active=true returns active only)
async function list(req, res) {
  try {
    const activeOnly = req.query.active === 'true';

    const result = await pool.query(
      activeOnly
        ? 'SELECT * FROM supervisors WHERE active = true ORDER BY id'
        : 'SELECT * FROM supervisors ORDER BY id'
    );

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

    if (!name || !String(name).trim()) {
      return res.status(400).json({ error: 'اسم المشرف مطلوب.' });
    }

    const result = await pool.query(
      'INSERT INTO supervisors(name, active) VALUES ($1,$2) RETURNING *',
      [String(name).trim(), active]
    );
    res.json({ message: 'Supervisor created successfully', supervisor: result.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server Error' });
  }
}

// Update an existing supervisor (partial update supported)
async function update(req, res) {
  try {
    const { id } = req.params;
    const { name, active } = req.body;

    const result = await pool.query(
      'UPDATE supervisors SET name=COALESCE($1,name), active=COALESCE($2,active) WHERE id=$3 RETURNING *',
      [name ?? null, active ?? null, id]
    );

    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'المشرف غير موجود.' });
    }

    res.json({ message: 'Supervisor updated successfully', supervisor: result.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server Error' });
  }
}

// Delete a supervisor
async function remove(req, res) {
  try {
    const { id } = req.params;

    const result = await pool.query(
      'DELETE FROM supervisors WHERE id=$1 RETURNING id',
      [id]
    );

    if (result.rowCount === 0) {
      return res.status(404).json({ error: 'المشرف غير موجود.' });
    }

    res.json({ success: true, message: 'Supervisor deleted successfully' });
  } catch (err) {
    // 23503 = foreign key violation (supervisor is used in plans/assignments)
    if (err.code === '23503') {
      return res.status(409).json({
        error: 'لا يمكن حذف المشرف لأنه مرتبط بخطط أو توزيعات سابقة. يمكنك إيقافه بدلًا من حذفه.'
      });
    }

    console.error(err);
    res.status(500).json({ error: 'Server Error' });
  }
}

module.exports = {
  list,
  create,
  update,
  remove
};