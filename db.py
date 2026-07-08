
# -*- coding: utf-8 -*-
import sqlite3
from datetime import datetime
from pathlib import Path
import pandas as pd

DB_PATH = "first_ops.db"


def connect(db_path=DB_PATH):
    conn = sqlite3.connect(db_path, check_same_thread=False, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def now():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def exec_sql(sql, params=(), db_path=DB_PATH):
    conn = connect(db_path)
    cur = conn.cursor()
    cur.execute(sql, params)
    conn.commit()
    last_id = cur.lastrowid
    conn.close()
    return last_id


def query_df(sql, params=(), db_path=DB_PATH):
    conn = connect(db_path)
    out = pd.read_sql_query(sql, conn, params=params)
    conn.close()
    return out


def init_db(db_path=DB_PATH):
    conn = connect(db_path)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        role TEXT DEFAULT 'Usuário',
        department TEXT,
        active INTEGER DEFAULT 1,
        created_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS routines (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL,
        description TEXT,
        department TEXT,
        owner TEXT,
        frequency TEXT DEFAULT 'Diária',
        priority TEXT DEFAULT 'Normal',
        mandatory INTEGER DEFAULT 0,
        start_date TEXT,
        project TEXT,
        active INTEGER DEFAULT 1,
        created_by TEXT,
        created_at TEXT,
        updated_at TEXT,
        UNIQUE(title, department, owner, project)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS executions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        routine_id INTEGER,
        execution_date TEXT,
        title TEXT,
        description TEXT,
        department TEXT,
        owner TEXT,
        project TEXT,
        priority TEXT,
        mandatory INTEGER,
        status TEXT DEFAULT 'Pendente',
        note TEXT,
        started_at TEXT,
        completed_by TEXT,
        completed_at TEXT,
        rescheduled_to TEXT,
        canceled_reason TEXT,
        created_at TEXT,
        updated_at TEXT,
        UNIQUE(routine_id, execution_date)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        description TEXT,
        department TEXT,
        owner TEXT,
        start_date TEXT,
        due_date TEXT,
        stage TEXT DEFAULT 'Planejamento',
        next_step TEXT,
        note TEXT,
        active INTEGER DEFAULT 1,
        created_at TEXT,
        updated_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        entity TEXT,
        entity_id INTEGER,
        event_date TEXT,
        user TEXT,
        action TEXT,
        note TEXT,
        created_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """)

    conn.commit()
    conn.close()


def setting(key, default=""):
    out = query_df("SELECT value FROM settings WHERE key=?", (key,))
    if out.empty:
        return default
    return str(out.iloc[0]["value"])


def set_setting(key, value):
    exec_sql("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(value)))


def log(entity, entity_id, event_date, user, action, note=""):
    exec_sql(
        "INSERT INTO events (entity, entity_id, event_date, user, action, note, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (entity, entity_id, event_date, user, action, note, now())
    )


def list_users(active=True):
    where = "WHERE active=1" if active else ""
    return query_df(f"SELECT * FROM users {where} ORDER BY name")


def create_user(name, role="Usuário", department=""):
    if not name:
        return None
    return exec_sql(
        "INSERT OR IGNORE INTO users (name, role, department, active, created_at) VALUES (?, ?, ?, 1, ?)",
        (name, role, department, now())
    )


def list_departments():
    out = query_df("""
    SELECT DISTINCT department FROM routines WHERE active=1 AND department IS NOT NULL AND TRIM(department) <> ''
    UNION
    SELECT DISTINCT department FROM users WHERE active=1 AND department IS NOT NULL AND TRIM(department) <> ''
    ORDER BY department
    """)
    return out["department"].dropna().astype(str).tolist() if not out.empty else []


def list_routines(active=True):
    where = "WHERE active=1" if active else ""
    return query_df(f"SELECT * FROM routines {where} ORDER BY department, owner, title")


def create_routine(data, user):
    routine_id = exec_sql(
        """
        INSERT OR IGNORE INTO routines (
            title, description, department, owner, frequency, priority, mandatory,
            start_date, project, active, created_by, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?)
        """,
        (
            data.get("title", ""),
            data.get("description", ""),
            data.get("department", ""),
            data.get("owner", ""),
            data.get("frequency", "Diária"),
            data.get("priority", "Normal"),
            int(data.get("mandatory", 0)),
            data.get("start_date", ""),
            data.get("project", ""),
            user,
            now(),
            now(),
        )
    )
    return routine_id


def archive_routine(routine_id, user):
    exec_sql("UPDATE routines SET active=0, updated_at=? WHERE id=?", (now(), routine_id))
    log("routine", routine_id, "", user, "Rotina arquivada", "")


def list_executions(execution_date=None, include_all=False):
    if execution_date and not include_all:
        return query_df("SELECT * FROM executions WHERE execution_date=? ORDER BY department, owner, title", (execution_date,))
    return query_df("SELECT * FROM executions ORDER BY execution_date DESC, department, owner, title")


def list_previous_pending(reference_date):
    return query_df(
        """
        SELECT * FROM executions
        WHERE execution_date < ?
          AND status NOT IN ('Concluída', 'Cancelada', 'Reprogramada')
        ORDER BY execution_date ASC, department, owner, title
        """,
        (reference_date,)
    )


def create_execution_from_routine(routine, execution_date):
    return exec_sql(
        """
        INSERT OR IGNORE INTO executions (
            routine_id, execution_date, title, description, department, owner, project,
            priority, mandatory, status, note, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', '', ?, ?)
        """,
        (
            int(routine["id"]),
            execution_date,
            routine.get("title", ""),
            routine.get("description", ""),
            routine.get("department", ""),
            routine.get("owner", ""),
            routine.get("project", ""),
            routine.get("priority", "Normal"),
            int(routine.get("mandatory", 0) or 0),
            now(),
            now(),
        )
    )


def set_execution_status(execution_id, status, user, note=""):
    if status == "Em andamento":
        exec_sql(
            "UPDATE executions SET status=?, note=?, started_at=?, updated_at=? WHERE id=?",
            (status, note, now(), now(), execution_id)
        )
    elif status == "Concluída":
        exec_sql(
            "UPDATE executions SET status=?, note=?, completed_by=?, completed_at=?, updated_at=? WHERE id=?",
            (status, note, user, now(), now(), execution_id)
        )
    else:
        exec_sql(
            "UPDATE executions SET status=?, note=?, updated_at=? WHERE id=?",
            (status, note, now(), execution_id)
        )

    ex = query_df("SELECT execution_date FROM executions WHERE id=?", (execution_id,))
    event_date = ex.iloc[0]["execution_date"] if not ex.empty else ""
    log("execution", execution_id, event_date, user, status, note)


def add_execution_note(execution_id, user, note):
    exec_sql("UPDATE executions SET note=?, updated_at=? WHERE id=?", (note, now(), execution_id))
    ex = query_df("SELECT execution_date FROM executions WHERE id=?", (execution_id,))
    event_date = ex.iloc[0]["execution_date"] if not ex.empty else ""
    log("execution", execution_id, event_date, user, "Observação", note)


def reschedule_execution(execution_id, user, new_date, note=""):
    ex = query_df("SELECT * FROM executions WHERE id=?", (execution_id,))
    if ex.empty:
        return
    row = ex.iloc[0].to_dict()
    exec_sql(
        "UPDATE executions SET status='Reprogramada', rescheduled_to=?, note=?, updated_at=? WHERE id=?",
        (new_date, note, now(), execution_id)
    )
    log("execution", execution_id, row.get("execution_date", ""), user, "Reprogramada", note)

    exec_sql(
        """
        INSERT OR IGNORE INTO executions (
            routine_id, execution_date, title, description, department, owner, project,
            priority, mandatory, status, note, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', ?, ?, ?)
        """,
        (
            row.get("routine_id"),
            new_date,
            row.get("title"),
            row.get("description"),
            row.get("department"),
            row.get("owner"),
            row.get("project"),
            row.get("priority"),
            row.get("mandatory"),
            f"Reprogramada de {row.get('execution_date')}: {note}",
            now(),
            now(),
        )
    )


def cancel_execution(execution_id, user, reason):
    exec_sql(
        "UPDATE executions SET status='Cancelada', canceled_reason=?, note=?, updated_at=? WHERE id=?",
        (reason, reason, now(), execution_id)
    )
    ex = query_df("SELECT execution_date FROM executions WHERE id=?", (execution_id,))
    event_date = ex.iloc[0]["execution_date"] if not ex.empty else ""
    log("execution", execution_id, event_date, user, "Cancelada", reason)


def list_projects(active=True):
    where = "WHERE active=1" if active else ""
    return query_df(f"SELECT * FROM projects {where} ORDER BY stage, due_date, name")


def create_project(data, user):
    pid = exec_sql(
        """
        INSERT OR IGNORE INTO projects (
            name, description, department, owner, start_date, due_date, stage,
            next_step, note, active, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
        """,
        (
            data.get("name", ""),
            data.get("description", ""),
            data.get("department", ""),
            data.get("owner", ""),
            data.get("start_date", ""),
            data.get("due_date", ""),
            data.get("stage", "Planejamento"),
            data.get("next_step", ""),
            data.get("note", ""),
            now(),
            now(),
        )
    )
    return pid


def list_events():
    return query_df("SELECT * FROM events ORDER BY id DESC")


def export_excel(path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        list_users(False).to_excel(writer, sheet_name="Usuarios", index=False)
        list_routines(False).to_excel(writer, sheet_name="Rotinas", index=False)
        list_executions(include_all=True).to_excel(writer, sheet_name="Execucoes", index=False)
        list_projects(False).to_excel(writer, sheet_name="Projetos", index=False)
        list_events().to_excel(writer, sheet_name="Historico", index=False)
