
# -*- coding: utf-8 -*-
from __future__ import annotations

import re
import shutil
import sqlite3
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
BACKUP_DIR = BASE_DIR / "backups"
DB_PATH = DATA_DIR / "first_ops_enterprise2.db"
GO_LIVE_DATE = date(2026, 7, 10)

DATA_DIR.mkdir(exist_ok=True)
BACKUP_DIR.mkdir(exist_ok=True)


def now_br() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def today_iso() -> str:
    return date.today().isoformat()


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def execute(sql: str, params: tuple = ()) -> int:
    with connect() as conn:
        cur = conn.execute(sql, params)
        conn.commit()
        return int(cur.lastrowid or 0)


def query(sql: str, params: tuple = ()) -> pd.DataFrame:
    with connect() as conn:
        return pd.read_sql_query(sql, conn, params=params)


def init_db() -> None:
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                role TEXT NOT NULL DEFAULT 'Usuário',
                department TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS routines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_id TEXT,
                title TEXT NOT NULL,
                description TEXT,
                department TEXT,
                owners TEXT NOT NULL,
                frequency TEXT NOT NULL DEFAULT 'Diária',
                due_rule TEXT,
                priority TEXT NOT NULL DEFAULT 'Normal',
                mandatory INTEGER NOT NULL DEFAULT 0,
                project TEXT,
                start_date TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_by TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(source_id)
            );

            CREATE TABLE IF NOT EXISTS daily_activities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                routine_id INTEGER NOT NULL,
                activity_date TEXT NOT NULL,
                title TEXT NOT NULL,
                description TEXT,
                department TEXT,
                owners TEXT NOT NULL,
                project TEXT,
                priority TEXT,
                due_time TEXT,
                status TEXT NOT NULL DEFAULT 'Pendente',
                note TEXT,
                started_by TEXT,
                started_at TEXT,
                completed_by TEXT,
                completed_at TEXT,
                rescheduled_to TEXT,
                canceled_by TEXT,
                canceled_at TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(routine_id, activity_date),
                FOREIGN KEY(routine_id) REFERENCES routines(id)
            );

            CREATE TABLE IF NOT EXISTS activity_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                activity_id INTEGER,
                routine_id INTEGER,
                event_date TEXT NOT NULL,
                event_time TEXT NOT NULL,
                user_name TEXT NOT NULL,
                action TEXT NOT NULL,
                note TEXT,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_id TEXT,
                name TEXT NOT NULL UNIQUE COLLATE NOCASE,
                objective TEXT,
                department TEXT,
                owners TEXT,
                start_date TEXT,
                due_date TEXT,
                status TEXT NOT NULL DEFAULT 'Em andamento',
                progress REAL NOT NULL DEFAULT 0,
                next_step TEXT,
                note TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            );
            """
        )
        conn.commit()


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def plain(value: Any) -> str:
    text = normalize_text(value).lower()
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


def normalize_department(value: Any) -> str:
    raw = normalize_text(value)
    lookup = {
        "contas a receber": "Contas a Receber",
        "contas receber": "Contas a Receber",
        "contas a pagar": "Contas a Pagar",
        "contas pagar": "Contas a Pagar",
        "contabilidade": "Contabilidade",
        "controladoria": "Controladoria",
        "tesouraria": "Tesouraria",
        "financeiro": "Financeiro",
        "coordenacao": "Coordenação",
        "projetos": "Projetos",
    }
    return lookup.get(plain(raw), raw.title())


def normalize_frequency(value: Any) -> str:
    key = plain(value)
    if key in {"diario", "diaria", "todos os dias", "todo dia", ""}:
        return "Diária"
    if key in {"mensal", "mes"}:
        return "Mensal"
    if key in {"quinzenal", "quinzena"}:
        return "Quinzenal"
    if key in {"semanal", "semana"}:
        return "Semanal"
    if key in {"unica", "unico", "pontual"}:
        return "Única"
    return normalize_text(value).title() or "Diária"


def normalize_priority(value: Any) -> str:
    key = plain(value)
    return {
        "critica": "Crítica",
        "critico": "Crítica",
        "alta": "Alta",
        "baixa": "Baixa",
        "normal": "Normal",
    }.get(key, normalize_text(value).title() or "Normal")


def owner_tokens(value: Any) -> list[str]:
    raw = normalize_text(value)
    parts = re.split(r"[/,;|]+", raw)
    return [part.strip() for part in parts if part.strip()]


def owner_matches(value: Any, user: str) -> bool:
    target = plain(user)
    return any(plain(name) == target for name in owner_tokens(value))


def get_setting(key: str, default: str = "") -> str:
    out = query("SELECT value FROM settings WHERE key=?", (key,))
    return default if out.empty else normalize_text(out.iloc[0]["value"])


def set_setting(key: str, value: Any) -> None:
    execute("INSERT OR REPLACE INTO settings(key, value) VALUES (?, ?)", (key, str(value)))


def list_users(active_only: bool = True) -> pd.DataFrame:
    where = "WHERE active=1" if active_only else ""
    return query(f"SELECT * FROM users {where} ORDER BY name")


def user_profile(name: str) -> dict:
    out = query("SELECT * FROM users WHERE LOWER(name)=LOWER(?) LIMIT 1", (name,))
    if out.empty:
        return {"name": name, "role": "Usuário", "department": ""}
    return out.iloc[0].to_dict()


def is_admin(name: str) -> bool:
    return plain(user_profile(name).get("role")) in {"administrador", "administradora", "admin"}


def create_user(name: str, role: str = "Usuário", department: str = "") -> None:
    if not normalize_text(name):
        return
    execute(
        """
        INSERT INTO users(name, role, department, active, created_at)
        VALUES (?, ?, ?, 1, ?)
        ON CONFLICT(name) DO UPDATE SET
            role=excluded.role,
            department=CASE
                WHEN TRIM(COALESCE(excluded.department,'')) <> '' THEN excluded.department
                ELSE users.department
            END,
            active=1
        """,
        (normalize_text(name), role or "Usuário", normalize_department(department), now_br()),
    )


def list_departments() -> list[str]:
    users = query("SELECT DISTINCT department FROM users WHERE active=1")
    routines = query("SELECT DISTINCT department FROM routines WHERE active=1")
    values = []
    for frame in (users, routines):
        if not frame.empty:
            values.extend(frame["department"].dropna().astype(str).tolist())
    return sorted({normalize_department(v) for v in values if normalize_text(v)})


def list_routines(active_only: bool = True) -> pd.DataFrame:
    where = "WHERE active=1" if active_only else ""
    return query(f"SELECT * FROM routines {where} ORDER BY department, owners, title, description")


def create_routine(data: dict, user: str) -> int:
    routine_id = execute(
        """
        INSERT INTO routines(
            source_id, title, description, department, owners, frequency,
            due_rule, priority, mandatory, project, start_date, active,
            created_by, created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, ?)
        ON CONFLICT(source_id)
        DO UPDATE SET
            title=excluded.title,
            description=excluded.description,
            department=excluded.department,
            owners=excluded.owners,
            frequency=excluded.frequency,
            due_rule=excluded.due_rule,
            priority=excluded.priority,
            mandatory=excluded.mandatory,
            project=excluded.project,
            start_date=excluded.start_date,
            active=1,
            updated_at=excluded.updated_at
        """,
        (
            normalize_text(data.get("source_id")),
            normalize_text(data.get("title")),
            normalize_text(data.get("description")),
            normalize_department(data.get("department")),
            normalize_text(data.get("owners")),
            normalize_frequency(data.get("frequency")),
            normalize_text(data.get("due_rule")),
            normalize_priority(data.get("priority")),
            int(bool(data.get("mandatory"))),
            normalize_text(data.get("project")),
            normalize_text(data.get("start_date")),
            user,
            now_br(),
            now_br(),
        ),
    )
    return routine_id


def update_routine(routine_id: int, data: dict, user: str) -> None:
    execute(
        """
        UPDATE routines SET
            title=?, description=?, department=?, owners=?, frequency=?,
            due_rule=?, priority=?, mandatory=?, project=?, start_date=?,
            active=?, updated_at=?
        WHERE id=?
        """,
        (
            normalize_text(data.get("title")),
            normalize_text(data.get("description")),
            normalize_department(data.get("department")),
            normalize_text(data.get("owners")),
            normalize_frequency(data.get("frequency")),
            normalize_text(data.get("due_rule")),
            normalize_priority(data.get("priority")),
            int(bool(data.get("mandatory"))),
            normalize_text(data.get("project")),
            normalize_text(data.get("start_date")),
            int(data.get("active", 1)),
            now_br(),
            routine_id,
        ),
    )


def list_projects(active_only: bool = True) -> pd.DataFrame:
    where = "WHERE active=1" if active_only else ""
    return query(f"SELECT * FROM projects {where} ORDER BY status, name")


def create_project(data: dict) -> None:
    execute(
        """
        INSERT INTO projects(
            source_id, name, objective, department, owners, start_date,
            due_date, status, progress, next_step, note, active,
            created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
        ON CONFLICT(name) DO UPDATE SET
            objective=excluded.objective,
            department=excluded.department,
            owners=excluded.owners,
            start_date=excluded.start_date,
            due_date=excluded.due_date,
            status=excluded.status,
            progress=excluded.progress,
            next_step=excluded.next_step,
            note=excluded.note,
            active=1,
            updated_at=excluded.updated_at
        """,
        (
            normalize_text(data.get("source_id")),
            normalize_text(data.get("name")),
            normalize_text(data.get("objective")),
            normalize_department(data.get("department")),
            normalize_text(data.get("owners")),
            normalize_text(data.get("start_date")),
            normalize_text(data.get("due_date")),
            normalize_text(data.get("status")) or "Em andamento",
            float(data.get("progress") or 0),
            normalize_text(data.get("next_step")),
            normalize_text(data.get("note")),
            now_br(),
            now_br(),
        ),
    )


def parse_date(value: Any) -> date | None:
    text = normalize_text(value)
    if not text:
        return None
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


def due_days(rule: Any) -> list[int]:
    text = normalize_text(rule)
    return [int(v) for v in re.findall(r"\b([0-3]?\d)\b", text) if 1 <= int(v) <= 31]


def due_time(rule: Any) -> str:
    text = normalize_text(rule)
    match = re.search(r"\b([01]?\d|2[0-3]):([0-5]\d)(?::[0-5]\d)?\b", text)
    return match.group(0)[:5] if match else ""


def is_business_day(day: date) -> bool:
    return day.weekday() < 5


def routine_occurs(row: pd.Series | dict, day: date) -> bool:
    if not is_business_day(day):
        return False

    start = parse_date(row.get("start_date"))
    if start and day < start:
        return False

    frequency = normalize_frequency(row.get("frequency"))

    if frequency == "Diária":
        return True

    if frequency == "Mensal":
        days = due_days(row.get("due_rule"))
        return day.day in (days or [10])

    if frequency == "Quinzenal":
        days = due_days(row.get("due_rule"))
        return day.day in (days or [5, 18])

    if frequency == "Semanal":
        base = start or date(2026, 1, 5)
        return day.weekday() == base.weekday()

    if frequency == "Única":
        return start == day

    return False


def ensure_activities(day: date) -> int:
    if not is_business_day(day):
        return 0

    routines = list_routines(True)
    created = 0

    for _, row in routines.iterrows():
        if not routine_occurs(row, day):
            continue

        before = query(
            "SELECT id FROM daily_activities WHERE routine_id=? AND activity_date=?",
            (int(row["id"]), day.isoformat()),
        )

        execute(
            """
            INSERT OR IGNORE INTO daily_activities(
                routine_id, activity_date, title, description, department,
                owners, project, priority, due_time, status, note,
                created_at, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', '', ?, ?)
            """,
            (
                int(row["id"]),
                day.isoformat(),
                row["title"],
                row.get("description", ""),
                row.get("department", ""),
                row.get("owners", ""),
                row.get("project", ""),
                row.get("priority", "Normal"),
                due_time(row.get("due_rule")),
                now_br(),
                now_br(),
            ),
        )

        if before.empty:
            created += 1

    return created


def backfill_until(day: date) -> None:
    """
    Materializa somente atividades a partir do Go Live.
    O período anterior é ignorado definitivamente.
    """
    if day < GO_LIVE_DATE:
        return

    last_text = get_setting("last_materialized_date", "")
    last = parse_date(last_text)

    if last is None or last < GO_LIVE_DATE:
        last = GO_LIVE_DATE - timedelta(days=1)

    cursor = last + timedelta(days=1)

    while cursor <= day:
        ensure_activities(cursor)
        cursor += timedelta(days=1)

    if day > last:
        set_setting("last_materialized_date", day.isoformat())


def list_activities(day: date | None = None) -> pd.DataFrame:
    if day is None:
        return query("SELECT * FROM daily_activities ORDER BY activity_date DESC, department, owners, title")
    ensure_activities(day)
    return query(
        "SELECT * FROM daily_activities WHERE activity_date=? ORDER BY department, owners, title, description",
        (day.isoformat(),),
    )


def list_user_activities(day: date, user: str) -> pd.DataFrame:
    frame = list_activities(day)
    if frame.empty:
        return frame
    return frame[frame["owners"].apply(lambda value: owner_matches(value, user))].copy()


def list_department_activities(day: date, department: str) -> pd.DataFrame:
    frame = list_activities(day)
    if frame.empty:
        return frame
    return frame[frame["department"].str.casefold() == department.casefold()].copy()


def list_previous_pending(reference_day: date, user: str | None = None) -> pd.DataFrame:
    """
    Considera pendência somente entre o Go Live e a data de referência.
    Atividades anteriores a 10/07/2026 são tratadas como implantação e não
    entram nos indicadores.
    """
    if reference_day <= GO_LIVE_DATE:
        return pd.DataFrame(columns=query(
            "SELECT * FROM daily_activities LIMIT 0"
        ).columns)

    frame = query(
        """
        SELECT * FROM daily_activities
        WHERE activity_date >= ?
          AND activity_date < ?
          AND status NOT IN ('Concluída', 'Cancelada', 'Reprogramada')
        ORDER BY activity_date, department, owners, title
        """,
        (GO_LIVE_DATE.isoformat(), reference_day.isoformat()),
    )

    if user and not frame.empty:
        frame = frame[
            frame["owners"].apply(lambda value: owner_matches(value, user))
        ].copy()

    return frame


def log_event(activity_id: int, routine_id: int, user: str, action: str, note: str = "") -> None:
    execute(
        """
        INSERT INTO activity_events(
            activity_id, routine_id, event_date, event_time,
            user_name, action, note, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            activity_id,
            routine_id,
            date.today().isoformat(),
            datetime.now().strftime("%H:%M:%S"),
            user,
            action,
            normalize_text(note),
            now_br(),
        ),
    )


def set_activity_status(activity_id: int, status: str, user: str, note: str = "") -> None:
    activity = query("SELECT * FROM daily_activities WHERE id=?", (activity_id,))
    if activity.empty:
        return
    row = activity.iloc[0]

    if status == "Em andamento":
        execute(
            """
            UPDATE daily_activities
            SET status=?, note=?, started_by=?, started_at=?, updated_at=?
            WHERE id=?
            """,
            (status, note, user, now_br(), now_br(), activity_id),
        )
    elif status == "Concluída":
        execute(
            """
            UPDATE daily_activities
            SET status=?, note=?, completed_by=?, completed_at=?, updated_at=?
            WHERE id=?
            """,
            (status, note, user, now_br(), now_br(), activity_id),
        )
    else:
        execute(
            "UPDATE daily_activities SET status=?, note=?, updated_at=? WHERE id=?",
            (status, note, now_br(), activity_id),
        )

    log_event(activity_id, int(row["routine_id"]), user, status, note)


def save_note(activity_id: int, user: str, note: str) -> None:
    activity = query("SELECT * FROM daily_activities WHERE id=?", (activity_id,))
    if activity.empty or not normalize_text(note):
        return
    row = activity.iloc[0]
    execute(
        "UPDATE daily_activities SET note=?, updated_at=? WHERE id=?",
        (note, now_br(), activity_id),
    )
    log_event(activity_id, int(row["routine_id"]), user, "Observação", note)


def reschedule_activity(activity_id: int, user: str, new_day: date, note: str) -> None:
    activity = query("SELECT * FROM daily_activities WHERE id=?", (activity_id,))
    if activity.empty:
        return
    row = activity.iloc[0]

    execute(
        """
        UPDATE daily_activities
        SET status='Reprogramada', rescheduled_to=?, note=?, updated_at=?
        WHERE id=?
        """,
        (new_day.isoformat(), note, now_br(), activity_id),
    )
    log_event(activity_id, int(row["routine_id"]), user, "Reprogramada", note)

    execute(
        """
        INSERT OR IGNORE INTO daily_activities(
            routine_id, activity_date, title, description, department,
            owners, project, priority, due_time, status, note,
            created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', ?, ?, ?)
        """,
        (
            int(row["routine_id"]),
            new_day.isoformat(),
            row["title"],
            row["description"],
            row["department"],
            row["owners"],
            row["project"],
            row["priority"],
            row["due_time"],
            f"Reprogramada de {row['activity_date']}. {note}".strip(),
            now_br(),
            now_br(),
        ),
    )


def cancel_activity(activity_id: int, user: str, note: str) -> None:
    activity = query("SELECT * FROM daily_activities WHERE id=?", (activity_id,))
    if activity.empty:
        return
    row = activity.iloc[0]
    execute(
        """
        UPDATE daily_activities
        SET status='Cancelada', note=?, canceled_by=?, canceled_at=?, updated_at=?
        WHERE id=?
        """,
        (note, user, now_br(), now_br(), activity_id),
    )
    log_event(activity_id, int(row["routine_id"]), user, "Cancelada", note)


def business_days_late(activity_date: str, reference_day: date) -> int:
    """
    Quantidade de dias úteis entre a data da atividade e a referência.
    Sexta para segunda representa 1 dia útil de atraso.
    """
    start = parse_date(activity_date)

    if start is None or reference_day <= start:
        return 0

    count = 0
    cursor = start + timedelta(days=1)

    while cursor <= reference_day:
        if is_business_day(cursor):
            count += 1
        cursor += timedelta(days=1)

    return count


def list_events(day: date | None = None) -> pd.DataFrame:
    if day is None:
        return query("SELECT * FROM activity_events ORDER BY id DESC")
    return query(
        "SELECT * FROM activity_events WHERE event_date=? ORDER BY id DESC",
        (day.isoformat(),),
    )


def daily_backup() -> Path:
    init_db()
    path = BACKUP_DIR / f"first_ops_{date.today().isoformat()}.db"
    if DB_PATH.exists() and not path.exists():
        shutil.copy2(DB_PATH, path)
    return path


def create_snapshot() -> Path:
    init_db()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = BACKUP_DIR / f"first_ops_snapshot_{stamp}.db"
    shutil.copy2(DB_PATH, path)
    return path



def list_backup_files() -> pd.DataFrame:
    rows = []
    for file in sorted(BACKUP_DIR.glob("*"), reverse=True):
        if file.is_file():
            stat = file.stat()
            rows.append({
                "arquivo": file.name,
                "tamanho_kb": round(stat.st_size / 1024, 1),
                "modificado_em": datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M:%S"),
                "caminho": str(file),
            })
    return pd.DataFrame(rows)


def cleanup_old_backups(keep_last: int = 30) -> int:
    files = sorted(
        [f for f in BACKUP_DIR.glob("*.db") if f.is_file()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    removed = 0
    for file in files[keep_last:]:
        try:
            file.unlink()
            removed += 1
        except OSError:
            pass
    return removed


def create_full_backup_package() -> Path:
    """
    Cria uma cópia instantânea do banco e uma exportação Excel.
    """
    snapshot = create_snapshot()
    excel_path = BACKUP_DIR / f"FIRST_OPS_Backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    export_excel(excel_path)
    return snapshot

def export_excel(path: Path) -> Path:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        list_users(False).to_excel(writer, "Usuarios", index=False)
        list_routines(False).to_excel(writer, "Rotinas", index=False)
        list_projects(False).to_excel(writer, "Projetos", index=False)
        list_activities().to_excel(writer, "Registro_Diario", index=False)
        list_events().to_excel(writer, "Historico", index=False)
    return path
