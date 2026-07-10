
# -*- coding: utf-8 -*-
from pathlib import Path
import unicodedata
import pandas as pd

from first_ops_database import (
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
            name = col(row, "Nome", "nome", "name", "Usuario", "Usuário", "usuario")
            if not name:
                continue
            role = col(row, "Perfil", "perfil", "role") or "Usuário"
            department = dep(col(row, "Departamento", "departamento", "department"))
            create_user(name, role, department)
            total_users += 1


    # Deriva usuários a partir dos responsáveis das tarefas, caso a aba Usuarios esteja incompleta
    if sh_tasks:
        try:
            tasks_for_users = pd.read_excel(excel_path, sheet_name=sh_tasks, dtype=object)
            for _, row in tasks_for_users.iterrows():
                resp = col(row, "Responsavel", "Responsável", "owner", "responsavel")
                department = dep(col(row, "Departamento", "departamento", "department"))
                if resp:
                    # aceita múltiplos responsáveis separados por /, ; ou ,
                    for name in str(resp).replace("/", ",").replace(";", ",").split(","):
                        name = name.strip()
                        if name:
                            create_user(name, "Usuário", department)
        except Exception:
            pass

    if sh_projects:
        projects_df = pd.read_excel(excel_path, sheet_name=sh_projects, dtype=object)
        for _, row in projects_df.iterrows():
            name = col(row, "Projeto", "projeto", "name")
            if not name:
                continue
            create_project({
                "name": name,
                "description": col(row, "Objetivo", "objetivo", "Descrição", "Descricao", "descricao", "description"),
                "department": dep(col(row, "Departamento", "departamento", "department")),
                "owner": col(row, "Responsavel", "Responsável", "responsavel", "owner"),
                "start_date": col(row, "Data de Inicio", "Data Início", "data_inicio", "start_date"),
                "due_date": col(row, "Prazo Final", "prazo_final", "due_date"),
                "stage": col(row, "Status", "status", "stage") or "Planejamento",
                "next_step": col(row, "Proxima Etapa", "Próxima Etapa", "proxima_etapa", "next_step"),
                "note": col(row, "Observação", "Observacao", "observacao", "note"),
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
                "description": col(row, "Descrição", "Descricao", "descricao", "description"),
                "department": dep(col(row, "Departamento", "departamento", "department")),
                "owner": col(row, "Responsavel", "Responsável", "responsavel", "owner"),
                "frequency": freq(col(row, "Periodicidade", "periodicidade", "frequency")),
                "priority": priority(col(row, "Prioridade", "prioridade", "priority")),
                "mandatory": yesno(col(row, "Obrigatoria", "Obrigatória", "obrigatoria", "mandatory")),
                "start_date": col(row, "Data de Inicio", "Data Início", "data_inicio", "start_date"),
                "project": col(row, "Projeto", "projeto", "project"),
            }, "Migração")
            total_routines += 1

    if sh_hist:
        hist_df = pd.read_excel(excel_path, sheet_name=sh_hist, dtype=object)
        conn = connect(DB_PATH)
        for _, row in hist_df.iterrows():
            action = col(row, "Status", "Ação", "Acao", "action")
            note = col(row, "Observação", "Observacao", "observacao", "note")
            user = col(row, "Usuário", "Usuario", "usuario", "user")
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


def repair_from_backup(excel_path="Agenda.xlsx"):
    """
    Reprocessa o backup e completa cadastros que entraram sem responsável
    ou sem classificação.
    """
    migrate_backup(excel_path, force=True)
    return repair_routines_from_backup(excel_path)


def _normalized_text(value):
    return norm(value)


def _find_existing_routine(title, description):
    routines = query_df(
        """
        SELECT *
        FROM routines
        WHERE LOWER(TRIM(title)) = LOWER(TRIM(?))
        ORDER BY
            CASE WHEN TRIM(COALESCE(owner, '')) = '' THEN 0 ELSE 1 END,
            id
        """,
        (title,)
    )

    if routines.empty:
        return None

    source_description = _normalized_text(description)

    if source_description:
        exact = routines[
            routines["description"].fillna("").apply(_normalized_text)
            == source_description
        ]
        if not exact.empty:
            return exact.iloc[0].to_dict()

    # Quando não há descrição, escolhe primeiro um cadastro ainda incompleto.
    incomplete = routines[
        (routines["owner"].fillna("").astype(str).str.strip() == "")
        | (routines["department"].fillna("").astype(str).str.strip() == "")
    ]
    if not incomplete.empty:
        return incomplete.iloc[0].to_dict()

    return routines.iloc[0].to_dict()


def repair_routines_from_backup(excel_path="Agenda.xlsx"):
    """
    Completa os cadastros importados anteriormente sem responsável,
    departamento, periodicidade, prioridade ou projeto.

    Não apaga execuções, histórico ou observações.
    """
    excel_path = Path(excel_path)

    if not excel_path.exists():
        return False, "Arquivo Agenda.xlsx não encontrado."

    xls = pd.ExcelFile(excel_path)
    sh_tasks = find_sheet(xls, ["Tarefas"])

    if not sh_tasks:
        return False, "A aba Tarefas não foi encontrada no backup."

    tasks_df = pd.read_excel(excel_path, sheet_name=sh_tasks, dtype=object)

    updated = 0
    created = 0
    users_created = 0

    for _, row in tasks_df.iterrows():
        title = col(row, "Tarefa", "tarefa", "title")
        if not title:
            continue

        description = col(
            row,
            "Descrição", "Descricao", "descricao", "description"
        )
        department = dep(
            col(row, "Departamento", "departamento", "department")
        )
        owner = col(
            row,
            "Responsavel", "Responsável", "responsavel", "owner"
        )
        frequency_value = freq(
            col(row, "Periodicidade", "periodicidade", "frequency")
        )
        priority_value = priority(
            col(row, "Prioridade", "prioridade", "priority")
        )
        mandatory_value = yesno(
            col(
                row,
                "Obrigatoria", "Obrigatória", "obrigatoria", "mandatory"
            )
        )
        start_date = col(
            row,
            "Data de Inicio", "Data Início", "data_inicio", "start_date"
        )
        project = col(row, "Projeto", "projeto", "project")

        # Garante que os responsáveis usados nas rotinas existam como usuários.
        if owner:
            for user_name in (
                str(owner)
                .replace("/", ",")
                .replace(";", ",")
                .replace("|", ",")
                .split(",")
            ):
                user_name = user_name.strip()
                if user_name:
                    create_user(user_name, "Usuário", department)
                    users_created += 1

        existing = _find_existing_routine(title, description)

        if existing:
            exec_sql(
                """
                UPDATE routines
                SET description=?,
                    department=?,
                    owner=?,
                    frequency=?,
                    priority=?,
                    mandatory=?,
                    start_date=?,
                    project=?,
                    active=1,
                    updated_at=?
                WHERE id=?
                """,
                (
                    description,
                    department,
                    owner,
                    frequency_value,
                    priority_value,
                    mandatory_value,
                    start_date,
                    project,
                    "",
                    int(existing["id"]),
                )
            )
            updated += 1
        else:
            create_routine(
                {
                    "title": title,
                    "description": description,
                    "department": department,
                    "owner": owner,
                    "frequency": frequency_value,
                    "priority": priority_value,
                    "mandatory": mandatory_value,
                    "start_date": start_date,
                    "project": project,
                },
                "Reparação do backup"
            )
            created += 1

    message = (
        f"Cadastros reparados: {updated} rotinas atualizadas, "
        f"{created} novas rotinas e responsáveis conferidos."
    )
    return True, message
