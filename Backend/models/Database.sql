CREATE EXTENSION IF NOT EXISTS "uuid-ossp";


-- =========================================================
-- SUPERVISORS
-- =========================================================

CREATE TABLE IF NOT EXISTS supervisors (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL,
    active BOOLEAN DEFAULT TRUE
);


-- =========================================================
-- PROFESSORS
-- =========================================================

CREATE TABLE IF NOT EXISTS professors (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL UNIQUE
);


-- =========================================================
-- COURSES
-- =========================================================

CREATE TABLE IF NOT EXISTS courses (
    id SERIAL PRIMARY KEY,
    crn TEXT,
    course_id TEXT,
    name TEXT
);


-- =========================================================
-- RAW SESSIONS
-- =========================================================

CREATE TABLE IF NOT EXISTS raw_sessions (
    id SERIAL PRIMARY KEY,

    excel_batch_id UUID NOT NULL,

    day_name TEXT,

    date DATE NOT NULL,

    period_label TEXT NOT NULL,

    time_from TIME NOT NULL,

    time_to TIME NOT NULL,

    sync_link TEXT,

    professor_name TEXT,

    course_text TEXT,

    crn TEXT,

    course_id TEXT,

    session_name TEXT,

    created_at TIMESTAMP DEFAULT NOW()
);


-- =========================================================
-- SESSION LINKS
-- =========================================================

CREATE TABLE IF NOT EXISTS session_links (
    id SERIAL PRIMARY KEY,

    raw_session_id INT
        REFERENCES raw_sessions(id)
        ON DELETE CASCADE,

    professor_id INT
        REFERENCES professors(id),

    course_id INT
        REFERENCES courses(id),

    crn TEXT,

    UNIQUE(raw_session_id)
);

CREATE TABLE IF NOT EXISTS room_assignments (
    id SERIAL PRIMARY KEY,

    plan_id UUID NOT NULL,

    professor_id INT
        REFERENCES professors(id),

    name TEXT,

    room_number TEXT NOT NULL,

    CONSTRAINT valid_room_number CHECK (
        room_number ~ '^[0-9]+$' AND (
            room_number::int BETWEEN 1 AND 14
            OR room_number::int BETWEEN 40 AND 46
        )
    ),

    UNIQUE (plan_id, professor_id)
);

CREATE INDEX IF NOT EXISTS
idx_room_assignments_plan
ON room_assignments(plan_id);

CREATE INDEX IF NOT EXISTS
idx_room_assignments_professor
ON room_assignments(professor_id);
-- =========================================================
-- SESSION GROUPS
-- =========================================================
CREATE TABLE IF NOT EXISTS session_groups (
    id SERIAL PRIMARY KEY,
    excel_batch_id UUID NOT NULL,
    date DATE NOT NULL,
    period_label TEXT NOT NULL,
    crn TEXT,
    professor_id INT REFERENCES professors(id),
    required_supervisors INT DEFAULT 1,
    sessions INT NOT NULL DEFAULT 1,
    time_from TEXT,
    time_to TEXT
);


-- =========================================================
-- SESSION GROUP INDEXES
-- =========================================================

CREATE INDEX IF NOT EXISTS
idx_session_groups_batch_date_period
ON session_groups (
    excel_batch_id,
    date,
    period_label
);

CREATE INDEX IF NOT EXISTS
idx_session_groups_crn
ON session_groups (crn);

CREATE INDEX IF NOT EXISTS
idx_session_groups_professor
ON session_groups (professor_id);


-- =========================================================
-- PLANS
-- =========================================================

CREATE TABLE IF NOT EXISTS plans (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),

    name TEXT,

    excel_batch_id UUID NOT NULL,

    date_from DATE NOT NULL,

    date_to DATE NOT NULL,

    created_at TIMESTAMP DEFAULT NOW()
);


-- =========================================================
-- DUTY POOL
-- =========================================================

CREATE TABLE IF NOT EXISTS duty_pool (
    id SERIAL PRIMARY KEY,

    plan_window_id UUID NOT NULL,

    supervisor_id INT
        REFERENCES supervisors(id)
        ON DELETE CASCADE,

    UNIQUE (
        plan_window_id,
        supervisor_id
    )
);

-- =========================================================
-- PREASSIGNMENTS
-- =========================================================

CREATE TABLE IF NOT EXISTS preassignments (
    id SERIAL PRIMARY KEY,

    plan_id UUID NOT NULL,

    session_group_id INT
        REFERENCES session_groups(id)
        ON DELETE CASCADE,

    supervisor_id INT
        REFERENCES supervisors(id)
        ON DELETE CASCADE,

    UNIQUE (
        plan_id,
        session_group_id,
        supervisor_id
    )
);


-- =========================================================
-- AFFINITIES
-- =========================================================

CREATE TABLE IF NOT EXISTS affinities (
    id SERIAL PRIMARY KEY,

    plan_id UUID NOT NULL,

    name TEXT,

    professor_id INT
        REFERENCES professors(id),

    crn TEXT,

    supervisor_id INT
        REFERENCES supervisors(id)
        ON DELETE CASCADE
);


-- =========================================================
-- ASSIGNMENT LOCKS
-- =========================================================

CREATE TABLE IF NOT EXISTS assignment_locks (
    id SERIAL PRIMARY KEY,

    plan_id UUID NOT NULL,

    session_group_id INT
        REFERENCES session_groups(id)
        ON DELETE CASCADE,

    supervisor_id INT
        REFERENCES supervisors(id)
        ON DELETE CASCADE,

    UNIQUE (
        plan_id,
        session_group_id
    )
);


-- =========================================================
-- ASSIGNMENTS
-- =========================================================

CREATE TABLE IF NOT EXISTS assignments (
    id SERIAL PRIMARY KEY,

    plan_id UUID NOT NULL,

    session_group_id INT
        REFERENCES session_groups(id)
        ON DELETE CASCADE,

    supervisor_id INT
        REFERENCES supervisors(id)
        ON DELETE CASCADE,

    UNIQUE (
        plan_id,
        session_group_id,
        supervisor_id
    )
);


-- =========================================================
-- INDEXES
-- =========================================================

CREATE INDEX IF NOT EXISTS
idx_raw_sessions_batch
ON raw_sessions(excel_batch_id);

CREATE INDEX IF NOT EXISTS
idx_raw_sessions_date_period
ON raw_sessions(date, period_label);

CREATE INDEX IF NOT EXISTS
idx_assignments_plan
ON assignments(plan_id);

CREATE INDEX IF NOT EXISTS
idx_assignments_supervisor
ON assignments(supervisor_id);

CREATE INDEX IF NOT EXISTS
idx_duty_pool_plan
ON duty_pool(plan_window_id);

-- ============================================================
-- Migration: rooms master table + room-conflict safety net
-- ============================================================

CREATE TABLE IF NOT EXISTS rooms (
  id SERIAL PRIMARY KEY,
  room_number VARCHAR(20) NOT NULL UNIQUE,
  capacity INTEGER,
  building VARCHAR(100),
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  created_at TIMESTAMP NOT NULL DEFAULT NOW()
);

INSERT INTO rooms (room_number) VALUES
  ('1'),('2'),('3'),('4'),('5'),('6'),('7'),('8'),('9'),('10'),
  ('11'),('12'),('13'),('14'),('15'),('16'),
  ('40'),('42'),('46'),('47'),('48'),('49'),
  ('100'),('101'),('102')
ON CONFLICT (room_number) DO NOTHING;

CREATE UNIQUE INDEX IF NOT EXISTS uq_room_assignments_room_slot
  ON room_assignments (plan_id, room_number, date, period_label)
  WHERE date IS NOT NULL AND period_label IS NOT NULL;