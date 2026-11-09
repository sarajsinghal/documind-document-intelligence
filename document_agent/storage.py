"""SQLite persistence for jobs, approved records, and append-only audit events."""
from __future__ import annotations
import json, sqlite3
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

def connect(path: str | Path) -> sqlite3.Connection:
    db = sqlite3.connect(path, check_same_thread=False)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys=ON")
    db.executescript("""
    CREATE TABLE IF NOT EXISTS jobs (
      id INTEGER PRIMARY KEY, filename TEXT NOT NULL, sha256 TEXT NOT NULL,
      doc_type TEXT NOT NULL, language TEXT NOT NULL, method TEXT NOT NULL,
      extracted_json TEXT NOT NULL, flags_json TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'pending',
      created_at TEXT NOT NULL, reviewed_at TEXT, reviewer_note TEXT NOT NULL DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS audit_events (
      id INTEGER PRIMARY KEY, job_id INTEGER NOT NULL REFERENCES jobs(id),
      event TEXT NOT NULL, details_json TEXT NOT NULL, created_at TEXT NOT NULL
    );
    """)
    db.commit()
    return db

def _now() -> str: return datetime.now(timezone.utc).isoformat(timespec="seconds")

def create_job(db: sqlite3.Connection, *, filename: str, sha256: str, doc_type: str, language: str,
               method: str, fields: dict[str, Any], flags: list[dict[str,str]]) -> int:
    cur = db.execute("INSERT INTO jobs(filename,sha256,doc_type,language,method,extracted_json,flags_json,created_at) VALUES(?,?,?,?,?,?,?,?)",
                     (filename,sha256,doc_type,language,method,json.dumps(fields,ensure_ascii=False),json.dumps(flags,ensure_ascii=False),_now()))
    job_id = int(cur.lastrowid)
    audit(db, job_id, "created", {"sha256": sha256, "method": method, "doc_type": doc_type})
    return job_id

def audit(db: sqlite3.Connection, job_id: int, event: str, details: dict[str, Any]) -> None:
    db.execute("INSERT INTO audit_events(job_id,event,details_json,created_at) VALUES(?,?,?,?)",
               (job_id,event,json.dumps(details,ensure_ascii=False),_now()))
    db.commit()

def review_job(db: sqlite3.Connection, job_id: int, fields: dict[str, Any], status: str, note: str) -> None:
    if status not in {"approved", "needs_changes", "pending"}: raise ValueError("Unsupported review state")
    before = db.execute("SELECT extracted_json,status FROM jobs WHERE id=?",(job_id,)).fetchone()
    if not before: raise ValueError("Job not found")
    encoded=json.dumps(fields,ensure_ascii=False)
    db.execute("UPDATE jobs SET extracted_json=?,status=?,reviewed_at=?,reviewer_note=? WHERE id=?",
               (encoded,status,_now(),note,job_id))
    audit(db,job_id,"reviewed",{"from_status":before["status"],"to_status":status,"fields_changed":before["extracted_json"] != encoded,"reviewer_note":note})
    db.commit()

def jobs(db: sqlite3.Connection) -> list[sqlite3.Row]:
    return list(db.execute("SELECT * FROM jobs ORDER BY id DESC"))

def events(db: sqlite3.Connection, job_id: int | None = None) -> list[sqlite3.Row]:
    if job_id is None: return list(db.execute("SELECT * FROM audit_events ORDER BY id DESC"))
    return list(db.execute("SELECT * FROM audit_events WHERE job_id=? ORDER BY id DESC",(job_id,)))
