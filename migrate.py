
# -*- coding: utf-8 -*-
from pathlib import Path
import unicodedata
import pandas as pd

from db import (
    init_db, connect, create_user, create_routine, create_project,
    exec_sql, query_df, set_setting, setting, log, DB_PATH
)


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def norm(v):
    t = clean(v).lower()
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode("ascii")
    return " ".join(t.split())


def col(row, *names):
    for n in names:
        if n in row and pd.notna(row[n]):
            return clean(row[n])
    return ""


def dep(v):
    mapa = {
        "contas a receber": "Contas a Receber",
        "contas receber": "Contas a Receber",
        "receber": "Contas a Receber",
        "contas a pagar": "Contas a Pagar",
        "contas pagar": "Contas a Pagar",
        "pagar": "Contas a Pagar",
        "contabilidade": "Contabilidade",
        "controladoria": "Controladoria",
        "tesouraria": "Tesouraria",
        "financeiro": "Financeiro",
        "projetos": "Projetos",
        "coordenacao": "Coordenação",
        "coordenação": "Coordenação",
    }
    return mapa.get(norm(v), clean(v).title())


def freq(v):
    mapa = {
        "diario": "Diária",
        "diaria": "Diária",
        "todo dia": "Diária",
        "todos os dias": "Diária",
        "semanal": "Semanal",
        "semana": "Semanal",
        "mensal": "Mensal",
        "mes": "Mensal",
        "unica": "Única",
        "unico": "Única",
        "pontual": "Única",
    }
    return mapa.get(norm(v), clean(v) or "Diária")


def priority(v):
    mapa = {
        "critica": "Crítica",
        "crítico": "Crítica",
        "critico": "Crítica",
        "alta": "Alta",
        "normal": "Normal",
        "baixa": "Baixa",
    }
    return mapa.get(norm(v), clean(v) or "Normal")


def yesno(v):
    return 1 if norm(v) in ["sim", "s", "1", "true", "obrigatoria", "obrigatório"] else 0


def find_sheet(xls, possibilities):
    low = {s.lower(): s for s in xls.sheet_names}
    for p in possibilities:
        if p.lower() in low:
            return low[p.lower()]
    return None


def preview_backup(excel_path="Agenda.xlsx"):
    excel_path = Path(excel_path)
    if not excel_path.exists():
        return {"found": False}

    xls = pd.ExcelFile(excel_path)
    sh_users = find_sheet(xls, ["Usuarios", "Usuários"])
    sh_tasks = find_sheet(xls, ["Tarefas"])
    sh_projects = find_sheet(xls, ["Projetos"])
    sh_hist = find_sheet(xls, ["Historico", "Histórico"])

    def rows(sheet):
        if not sheet:
            return 0
        try:
            return len(pd.read_excel(excel_path, sheet_name=sheet, dtype=object))
        except Exception:
            return 0

    return {
        "found": True,
        "file": str(excel_path),
        "users": rows(sh_users),
        "tasks": rows(sh_tasks),
        "projects": rows(sh_projects),
        "history": rows(sh_hist),
        "sheets": xls.sheet_names,
    }


def migrate_backup(excel_path="Agenda.xlsx", force=False):
    excel_path = Path(excel_path)
    if not excel_path.exists():
        return False, "Backup Agenda.xlsx não encontrado."

    init_db(DB_PATH)

    if setting("migration_done", "0") == "1" and not force:
        return False, "Migração já realizada. Para repetir, use a opção forçada na Administração."

    xls = pd.ExcelFile(excel_path)
    sh_users = find_sheet(xls, ["Usuarios", "Usuários"])
    sh_tasks = find_sheet(xls, ["Tarefas"])
    sh_projects = find_sheet(xls, ["Projetos"])
    sh_hist = find_sheet(xls, ["Historico", "Histórico"])

    total_users = total_routines = total_projects = total_hist = 0

    if sh_users:
        users_df = pd.read_excel(excel_path, sheet_name=sh_users, dtype=object)
        for _, row in users_df.iterrows():
            name = col(row, "Nome", "name", "Usuario", "Usuário")
            if not name:
                continue
            role = col(row, "Perfil", "role") or "Usuário"
            department = dep(col(row, "Departamento", "department"))
            create_user(name, role, department)
            total_users += 1

    if sh_projects:
        projects_df = pd.read_excel(excel_path, sheet_name=sh_projects, dtype=object)
        for _, row in projects_df.iterrows():
            name = col(row, "Projeto", "projeto", "name")
            if not name:
                continue
            create_project({
                "name": name,
                "description": col(row, "Objetivo", "Descrição", "Descricao", "description"),
                "department": dep(col(row, "Departamento", "department")),
                "owner": col(row, "Responsavel", "Responsável", "owner"),
                "start_date": col(row, "Data de Inicio", "Data Início", "data_inicio", "start_date"),
                "due_date": col(row, "Prazo Final", "prazo_final", "due_date"),
                "stage": col(row, "Status", "stage") or "Planejamento",
                "next_step": col(row, "Proxima Etapa", "Próxima Etapa", "next_step"),
                "note": col(row, "Observação", "Observacao", "note"),
            }, "Migração")
            total_projects += 1

    if sh_tasks:
        tasks_df = pd.read_excel(excel_path, sheet_name=sh_tasks, dtype=object)
        for _, row in tasks_df.iterrows():
            title = col(row, "Tarefa", "tarefa", "title")
            if not title:
                continue

            create_routine({
                "title": title,
                "description": col(row, "Descrição", "Descricao", "description"),
                "department": dep(col(row, "Departamento", "department")),
                "owner": col(row, "Responsavel", "Responsável", "owner"),
                "frequency": freq(col(row, "Periodicidade", "frequency")),
                "priority": priority(col(row, "Prioridade", "priority")),
                "mandatory": yesno(col(row, "Obrigatoria", "Obrigatória", "mandatory")),
                "start_date": col(row, "Data de Inicio", "Data Início", "data_inicio", "start_date"),
                "project": col(row, "Projeto", "project"),
            }, "Migração")
            total_routines += 1

    if sh_hist:
        hist_df = pd.read_excel(excel_path, sheet_name=sh_hist, dtype=object)
        conn = connect(DB_PATH)
        for _, row in hist_df.iterrows():
            action = col(row, "Status", "Ação", "Acao", "action")
            note = col(row, "Observação", "Observacao", "note")
            user = col(row, "Usuário", "Usuario", "user")
            data = col(row, "Data", "data", "event_date")
            tarefa = col(row, "Tarefa", "tarefa")
            conn.execute(
                """
                INSERT INTO events (entity, entity_id, event_date, user, action, note, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                ("legacy", None, data, user or "Migração", action or "Histórico legado", f"{tarefa} | {note}", "")
            )
            total_hist += 1
        conn.commit()
        conn.close()

    set_setting("migration_done", "1")
    set_setting("migration_file", str(excel_path))

    msg = f"Migração concluída: {total_users} usuários, {total_routines} rotinas, {total_projects} projetos, {total_hist} históricos."
    return True, msg
