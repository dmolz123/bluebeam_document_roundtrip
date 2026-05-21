/**
 * db.js — SQLite persistence layer for Atkins QA/QC PoC
 *
 * DB file location is controlled by DB_PATH env var:
 *   - Local dev:  DB_PATH=./poc.db  (add to .env)
 *   - Render:     DB_PATH=/data/poc.db  (set in Render Environment tab)
 *
 * Schema:
 *   projects   — one row per Atkins Project ID
 *   sessions   — one row per Bluebeam Studio Session
 *   files      — one row per file uploaded/checked out
 *   snapshots  — markup metadata snapshots after each downstream run
 */

const Database = require('better-sqlite3');
const path = require('path');

const DB_PATH = process.env.DB_PATH || './poc.db';

console.log(`[DB] Opening database at: ${DB_PATH}`);

const db = new Database(DB_PATH);

// WAL mode = better read concurrency, safer on Render disk
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

// ---------------------------------------------------------------------------
// SCHEMA — safe to run on every startup (IF NOT EXISTS guards)
// ---------------------------------------------------------------------------
db.exec(`
  CREATE TABLE IF NOT EXISTS projects (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    atkins_project_id   TEXT    NOT NULL UNIQUE,
    bluebeam_project_id TEXT,
    project_name        TEXT,
    region              TEXT,
    review_type         TEXT,
    qa_category         TEXT,
    created_at          TEXT    NOT NULL DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS sessions (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    atkins_project_id     TEXT    NOT NULL,
    bluebeam_session_id   TEXT    NOT NULL UNIQUE,
    review_name           TEXT,
    discipline            TEXT,
    review_type           TEXT,
    polling_interval      INTEGER NOT NULL DEFAULT 0,
    status                TEXT    NOT NULL DEFAULT 'active',
    created_at            TEXT    NOT NULL DEFAULT (datetime('now')),
    last_polled_at        TEXT
  );

  CREATE TABLE IF NOT EXISTS files (
    id                        INTEGER PRIMARY KEY AUTOINCREMENT,
    atkins_project_id         TEXT    NOT NULL,
    bluebeam_project_id       TEXT,
    bluebeam_project_file_id  TEXT,
    bluebeam_session_id       TEXT,
    bluebeam_session_file_id  TEXT,
    file_name                 TEXT,
    file_size                 INTEGER,
    created_at                TEXT    NOT NULL DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS snapshots (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    atkins_project_id     TEXT    NOT NULL,
    bluebeam_session_id   TEXT    NOT NULL,
    markup_count          INTEGER NOT NULL DEFAULT 0,
    snapshot_json         TEXT,
    created_at            TEXT    NOT NULL DEFAULT (datetime('now'))
  );

  CREATE INDEX IF NOT EXISTS idx_sessions_atkins_id
    ON sessions(atkins_project_id);

  CREATE INDEX IF NOT EXISTS idx_files_atkins_id
    ON files(atkins_project_id);

  CREATE INDEX IF NOT EXISTS idx_snapshots_atkins_session
    ON snapshots(atkins_project_id, bluebeam_session_id);
`);

console.log('[DB] Schema ready');

// ---------------------------------------------------------------------------
// PREPARED STATEMENTS
// ---------------------------------------------------------------------------

// projects
const upsertProject = db.prepare(`
  INSERT INTO projects (atkins_project_id, bluebeam_project_id, project_name, region, review_type, qa_category)
  VALUES (@atkinsProjectId, @bluebeamProjectId, @projectName, @region, @reviewType, @qaCategory)
  ON CONFLICT(atkins_project_id) DO UPDATE SET
    bluebeam_project_id = excluded.bluebeam_project_id,
    project_name        = excluded.project_name,
    region              = excluded.region,
    review_type         = excluded.review_type,
    qa_category         = excluded.qa_category
`);

const getProject = db.prepare(`
  SELECT * FROM projects WHERE atkins_project_id = ?
`);

const listProjects = db.prepare(`
  SELECT * FROM projects ORDER BY created_at DESC
`);

// sessions
const insertSession = db.prepare(`
  INSERT OR IGNORE INTO sessions
    (atkins_project_id, bluebeam_session_id, review_name, discipline, review_type, polling_interval)
  VALUES
    (@atkinsProjectId, @bluebeamSessionId, @reviewName, @discipline, @reviewType, @pollingInterval)
`);

const updateSessionPolled = db.prepare(`
  UPDATE sessions
  SET last_polled_at = datetime('now')
  WHERE bluebeam_session_id = ?
`);

const updateSessionStatus = db.prepare(`
  UPDATE sessions SET status = ? WHERE bluebeam_session_id = ?
`);

const getSessionsByProject = db.prepare(`
  SELECT * FROM sessions WHERE atkins_project_id = ? ORDER BY created_at DESC
`);

// files
const insertFile = db.prepare(`
  INSERT INTO files
    (atkins_project_id, bluebeam_project_id, bluebeam_project_file_id, bluebeam_session_id, bluebeam_session_file_id, file_name, file_size)
  VALUES
    (@atkinsProjectId, @bluebeamProjectId, @bluebeamProjectFileId, @bluebeamSessionId, @bluebeamSessionFileId, @fileName, @fileSize)
`);

const updateFileSession = db.prepare(`
  UPDATE files
  SET bluebeam_session_id = @sessionId, bluebeam_session_file_id = @sessionFileId
  WHERE bluebeam_project_file_id = @projectFileId
`);

const getFilesByProject = db.prepare(`
  SELECT * FROM files WHERE atkins_project_id = ? ORDER BY created_at DESC
`);

const getFilesBySession = db.prepare(`
  SELECT * FROM files WHERE bluebeam_session_id = ? ORDER BY created_at DESC
`);

// snapshots
const insertSnapshot = db.prepare(`
  INSERT INTO snapshots (atkins_project_id, bluebeam_session_id, markup_count, snapshot_json)
  VALUES (@atkinsProjectId, @bluebeamSessionId, @markupCount, @snapshotJson)
`);

const getLatestSnapshotBySession = db.prepare(`
  SELECT * FROM snapshots
  WHERE atkins_project_id = ? AND bluebeam_session_id = ?
  ORDER BY created_at DESC
  LIMIT 1
`);

const getSnapshotsByProject = db.prepare(`
  SELECT * FROM snapshots
  WHERE atkins_project_id = ?
  ORDER BY created_at DESC
`);

// ---------------------------------------------------------------------------
// EXPORTED API
// ---------------------------------------------------------------------------
module.exports = {
  // ── projects ──────────────────────────────────────────────────────────────
  upsertProject: (params) => upsertProject.run(params),
  getProject:    (atkinsProjectId) => getProject.get(atkinsProjectId),
  listProjects:  () => listProjects.all(),

  // ── sessions ──────────────────────────────────────────────────────────────
  insertSession:        (params) => insertSession.run(params),
  updateSessionPolled:  (sessionId) => updateSessionPolled.run(sessionId),
  updateSessionStatus:  (sessionId, status) => updateSessionStatus.run(status, sessionId),
  getSessionsByProject: (atkinsProjectId) => getSessionsByProject.all(atkinsProjectId),

  // ── files ─────────────────────────────────────────────────────────────────
  insertFile:       (params) => insertFile.run(params),
  updateFileSession:(params) => updateFileSession.run(params),
  getFilesByProject:(atkinsProjectId) => getFilesByProject.all(atkinsProjectId),
  getFilesBySession:(sessionId) => getFilesBySession.all(sessionId),

  // ── snapshots ─────────────────────────────────────────────────────────────
  insertSnapshot:           (params) => insertSnapshot.run(params),
  getLatestSnapshotBySession:(atkinsProjectId, sessionId) =>
    getLatestSnapshotBySession.get(atkinsProjectId, sessionId),
  getSnapshotsByProject:    (atkinsProjectId) => getSnapshotsByProject.all(atkinsProjectId),

  // ── raw db access (for transactions) ──────────────────────────────────────
  db,
};
