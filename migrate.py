# -*- coding: utf-8 -*-
from pathlib import Path
from typing import Any

import pandas as pd

from first_ops_database_v232 import (
    DB_PATH,
    connect,
    create_project,
    create_routine,
    create_user,
    init_db,
    normalize_department,
    normalize_text,
    set_setting,
)


def col(row: pd.Series, *names: str) -> str:
    for name in names:
        if name in row.index and pd.notna(row[name]):
            return normalize_text(row[name])
    return ""


def yes(value: Any) -> bool:
    return normalize_text(value).lower() in {"sim", "s", "1", "true"}


def database_has_master_data() -> bool:
    with connect() as conn:
        users = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        routines = conn.execute("SELECT COUNT(*) FROM routines").fetchone()[0]
    return users > 0 and routines > 0


def import_backup(excel_path: str = "Agenda.xlsx", force: bool = False) -> tuple[bool, str]:
    source = Path(excel_path)
    if not source.exists():
        return False, "Arquivo Agenda.xlsx não encontrado."

    init_db()

    if database_has_master_data() and not force:
        return True, "Cadastros já carregados."

    xls = pd.ExcelFile(source)
    users = pd.read_excel(source, "Usuarios", dtype=object)
    tasks = pd.read_excel(source, "Tarefas", dtype=object)
    projects = pd.read_excel(source, "Projetos", dtype=object)
    history = pd.read_excel(source, "Historico", dtype=object)

    user_count = 0
    for _, row in users.iterrows():
        name = col(row, "nome", "Nome")
        if not name:
            continue
        role = col(row, "perfil", "Perfil") or "Usuário"
        department = col(row, "departamento", "Departamento")
        create_user(name, role, department)
        user_count += 1

    routine_count = 0
    for _, row in tasks.iterrows():
        title = col(row, "tarefa", "Tarefa")
        owners = col(row, "responsavel", "Responsável", "Responsavel")
        if not title or not owners:
            continue

        create_routine(
            {
                "source_id": col(row, "id"),
                "title": title,
                "description": col(row, "descricao", "Descrição", "Descricao"),
                "department": col(row, "departamento", "Departamento"),
                "owners": owners,
                "frequency": col(row, "periodicidade", "Periodicidade") or "Diária",
                "due_rule": col(row, "prazo_limite", "Prazo Limite"),
                "priority": col(row, "prioridade", "Prioridade") or "Normal",
                "mandatory": yes(col(row, "obrigatoria", "Obrigatória", "Obrigatoria")),
                "project": col(row, "projeto", "Projeto"),
                "start_date": col(row, "data_inicio", "Data Início"),
            },
            col(row, "criado_por", "Criado Por") or "Migração",
        )

        for owner in owners.replace("/", ",").replace(";", ",").split(","):
            owner = owner.strip()
            if owner:
                create_user(owner, "Usuário", col(row, "departamento", "Departamento"))

        routine_count += 1

    project_count = 0
    for _, row in projects.iterrows():
        name = col(row, "projeto", "Projeto")
        if not name:
            continue
        create_project(
            {
                "source_id": col(row, "id"),
                "name": name,
                "objective": col(row, "objetivo", "Objetivo"),
                "department": col(row, "departamento", "Departamento"),
                "owners": col(row, "responsavel", "Responsável", "Responsavel"),
                "start_date": col(row, "data_inicio", "Data Início"),
                "due_date": col(row, "prazo_final", "Prazo Final"),
                "status": col(row, "status", "Status") or "Em andamento",
                "progress": col(row, "percentual", "Percentual") or 0,
                "next_step": col(row, "proxima_etapa", "Próxima Etapa"),
                "note": col(row, "observacao", "Observação"),
            }
        )
        project_count += 1

    legacy_count = 0
    with connect() as conn:
        for _, row in history.iterrows():
            event_date = pd.to_datetime(
                col(row, "data", "Data"), dayfirst=True, errors="coerce"
            )
            if pd.isna(event_date):
                continue
            conn.execute(
                """
                INSERT INTO activity_events(
                    activity_id, routine_id, event_date, event_time,
                    user_name, action, note, created_at
                )
                VALUES (NULL, NULL, ?, ?, ?, ?, ?, ?)
                """,
                (
                    event_date.date().isoformat(),
                    col(row, "hora", "Hora") or "00:00:00",
                    col(row, "usuario", "Usuário") or "Migração",
                    col(row, "status", "Status") or "Histórico legado",
                    f"{col(row, 'tarefa', 'Tarefa')} | {col(row, 'observacao', 'Observação')}",
                    col(row, "criado_em", "Criado em"),
                ),
            )
            legacy_count += 1
        conn.commit()

    set_setting("backup_imported", "1")

    return True, (
        f"Importação concluída: {user_count} usuários, "
        f"{routine_count} rotinas, {project_count} projetos e "
        f"{legacy_count} registros históricos."
    )
