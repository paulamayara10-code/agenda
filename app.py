
# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
from io import BytesIO
import unicodedata

import pandas as pd
import streamlit as st

from first_ops_database_v232 import (
    BACKUP_DIR,
    GO_LIVE_DATE,
    DB_PATH,
    backfill_until,
    cancel_activity,
    create_project,
    create_routine,
    create_snapshot,
    create_full_backup_package,
    list_backup_files,
    cleanup_old_backups,
    create_user,
    daily_backup,
    ensure_activities,
    export_excel,
    is_admin,
    is_business_day,
    list_activities,
    list_department_activities,
    list_departments,
    list_events,
    list_previous_pending,
    list_projects,
    list_routines,
    list_user_activities,
    list_users,
    normalize_department,
    normalize_text,
    owner_matches,
    reschedule_activity,
    save_note,
    set_activity_status,
    update_routine,
    get_routine,
    set_routine_active,
    duplicate_routine,
    delete_routine_permanently,
    user_profile,
)
from migrate import import_backup


st.set_page_config(
    page_title="FIRST OPS Enterprise 2.3.4",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Banco novo e isolado.
import_backup("Agenda.xlsx", force=False)
backfill_until(date.today())
daily_backup()


st.markdown(
    """
    <style>
    .stApp { background:#f4f7fb; }
    div.block-container { padding-top:1.2rem; padding-bottom:4rem; max-width:1500px; }
    section[data-testid="stSidebar"] {
        background:linear-gradient(180deg,#071020 0%,#111827 58%,#312e81 100%);
    }
    section[data-testid="stSidebar"] * { color:#ffffff !important; }
    section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
        background:#ffffff !important;
        border:1px solid #cbd5e1 !important;
        border-radius:12px !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="select"] input,
    section[data-testid="stSidebar"] div[data-baseweb="select"] span,
    section[data-testid="stSidebar"] div[data-baseweb="select"] div {
        color:#0f172a !important;
        -webkit-text-fill-color:#0f172a !important;
        font-weight:700 !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="select"] svg {
        fill:#0f172a !important;
        color:#0f172a !important;
    }
    .ops-title { font-size:38px; font-weight:950; color:#0f172a; letter-spacing:-.04em; }
    .ops-sub { color:#64748b; margin:3px 0 18px 0; }
    .hero {
        background:linear-gradient(135deg,#111827,#312e81 58%,#6d28d9);
        color:#fff; border-radius:28px; padding:27px 30px;
        box-shadow:0 20px 55px rgba(49,46,129,.24); margin-bottom:20px;
    }
    .hero h1 { margin:0; font-size:31px; font-weight:950; }
    .hero p { margin:8px 0 0 0; color:#ddd6fe; }
    .panel {
        background:#fff; border:1px solid #e2e8f0; border-radius:24px;
        padding:20px; box-shadow:0 14px 34px rgba(15,23,42,.06); margin-top:16px;
    }
    .panel-title { font-size:21px; font-weight:900; color:#0f172a; margin-bottom:14px; }
    .metric {
        background:#fff; border:1px solid #e2e8f0; border-radius:22px;
        padding:20px; min-height:118px; box-shadow:0 13px 30px rgba(15,23,42,.055);
    }
    .metric-label { color:#64748b; text-transform:uppercase; font-size:12px; font-weight:850; }
    .metric-value { font-size:35px; font-weight:950; margin-top:7px; }
    .task-title { font-size:18px; font-weight:900; color:#0f172a; }
    .task-done { font-size:18px; font-weight:850; color:#94a3b8; text-decoration:line-through; }
    .tag {
        display:inline-block; padding:5px 9px; border-radius:999px;
        font-size:12px; font-weight:850; margin:7px 5px 0 0;
    }
    .blue{background:#dbeafe;color:#1d4ed8}.green{background:#dcfce7;color:#15803d}
    .red{background:#fee2e2;color:#b91c1c}.yellow{background:#fef3c7;color:#b45309}
    .purple{background:#ede9fe;color:#6d28d9}.gray{background:#f1f5f9;color:#475569}
    .orange{background:#ffedd5;color:#c2410c}
    .stButton>button { border-radius:13px; font-weight:800; }
    </style>
    """,
    unsafe_allow_html=True,
)


def plain(value: str) -> str:
    text = str(value or "").strip().lower()
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


def br(day: date) -> str:
    return day.strftime("%d/%m/%Y")


def ref_date() -> date:
    if "reference_date" not in st.session_state:
        st.session_state["reference_date"] = date.today()
    return st.session_state["reference_date"]


def move_date(days: int) -> None:
    st.session_state["reference_date"] = ref_date() + timedelta(days=days)


def go_today() -> None:
    st.session_state["reference_date"] = date.today()


def header(title: str, subtitle: str) -> None:
    st.markdown(f"<div class='ops-title'>{title}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='ops-sub'>{subtitle}</div>", unsafe_allow_html=True)


def metric(label: str, value: int | str, detail: str, color: str) -> None:
    st.markdown(
        f"""
        <div class='metric'>
            <div class='metric-label'>{label}</div>
            <div class='metric-value' style='color:{color}'>{value}</div>
            <div style='color:{color};font-size:13px'>{detail}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def date_bar() -> date:
    st.markdown("<div class='panel' style='padding:14px 17px'>", unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns([1.15, 1.1, .85, .65, 2.5])

    with c1:
        st.markdown("#### 📅 Data de referência")
    with c2:
        selected = st.date_input(
            "Data",
            value=ref_date(),
            label_visibility="collapsed",
            key="global_reference_date",
        )
        st.session_state["reference_date"] = selected
    with c3:
        st.button("◀ Anterior", use_container_width=True, on_click=move_date, args=(-1,))
    with c4:
        st.button("Hoje", use_container_width=True, on_click=go_today)
    with c5:
        a, b = st.columns([.8, 1.7])
        with a:
            st.button("Próximo ▶", use_container_width=True, on_click=move_date, args=(1,))
        with b:
            if is_business_day(ref_date()):
                st.success(f"Dia útil: {br(ref_date())}")
            else:
                st.warning("Data sem rotina automática.")

    st.markdown("</div>", unsafe_allow_html=True)
    return ref_date()


def frame(title: str, subtitle: str) -> date:
    header(title, subtitle)
    return date_bar()


def sidebar():
    users = list_users(True)
    names = users["name"].dropna().astype(str).tolist()

    if not names:
        create_user("Paula", "Administradora", "Controladoria")
        names = ["Paula"]

    st.sidebar.markdown("## ✅ FIRST OPS")
    st.sidebar.markdown("### Gestão Operacional")
    st.sidebar.divider()

    user = st.sidebar.selectbox("Usuário", names)
    profile = user_profile(user)

    st.sidebar.caption(
        f"{profile.get('role') or 'Usuário'} • "
        f"{profile.get('department') or 'Sem departamento'}"
    )

    fixed = {
        "inicio", "meu dia", "equipe", "coordenacao", "pendencias",
        "rotinas", "projetos", "historico", "administracao",
        "base e backup",
    }
    departments = [
        item for item in list_departments()
        if plain(item) not in fixed
    ]

    menu = ["Meu Painel", "Meu Dia", "Equipe", "Coordenação"]
    menu += departments
    menu += ["Pendências", "Rotinas", "Projetos", "Histórico", "Usuários", "Backup"]

    page = st.sidebar.radio("Navegação", menu)
    return user, page, departments


def status_class(status: str) -> str:
    return {
        "Pendente": "yellow",
        "Em andamento": "orange",
        "Concluída": "green",
        "Cancelada": "gray",
        "Reprogramada": "blue",
    }.get(status, "purple")


def activity_card(row: pd.Series, user: str, prefix: str) -> None:
    activity_id = int(row["id"])
    status = normalize_text(row.get("status")) or "Pendente"
    title = normalize_text(row.get("title"))
    description = normalize_text(row.get("description"))

    with st.container(border=True):
        left, right = st.columns([5.2, 1.15])

        with left:
            checked = status == "Concluída"
            new_checked = st.checkbox(
                title,
                value=checked,
                key=f"check_{prefix}_{activity_id}",
                disabled=checked,
            )

            if new_checked and not checked:
                set_activity_status(
                    activity_id,
                    "Concluída",
                    user,
                    "Concluída pelo checklist rápido",
                )
                st.rerun()

            if description:
                st.caption(description)

            tags = f"<span class='tag {status_class(status)}'>{status}</span>"
            tags += f"<span class='tag blue'>{normalize_text(row.get('department'))}</span>"
            tags += f"<span class='tag green'>👤 {normalize_text(row.get('owners'))}</span>"

            if normalize_text(row.get("project")):
                tags += f"<span class='tag purple'>📁 {normalize_text(row.get('project'))}</span>"

            if normalize_text(row.get("priority")):
                tags += f"<span class='tag yellow'>{normalize_text(row.get('priority'))}</span>"

            if normalize_text(row.get("due_time")):
                tags += f"<span class='tag red'>⏰ {normalize_text(row.get('due_time'))}</span>"

            st.markdown(tags, unsafe_allow_html=True)

            if normalize_text(row.get("note")):
                st.info(normalize_text(row.get("note")))

        with right:
            if status == "Concluída":
                st.success("Concluída")
            elif status == "Em andamento":
                st.warning("Em andamento")
            else:
                if st.button("▶ Iniciar", key=f"start_{prefix}_{activity_id}", use_container_width=True):
                    set_activity_status(
                        activity_id,
                        "Em andamento",
                        user,
                        "Atividade iniciada",
                    )
                    st.rerun()

        with st.expander("💬 Detalhes"):
            note = st.text_area(
                "Observação",
                key=f"note_{prefix}_{activity_id}",
                placeholder="Ex.: aguardando retorno do banco, documento não recebido...",
            )

            a, b, c = st.columns(3)

            with a:
                if st.button("Salvar observação", key=f"save_note_{prefix}_{activity_id}"):
                    if note.strip():
                        save_note(activity_id, user, note)
                        st.rerun()
                    st.warning("Digite uma observação.")

            with b:
                new_day = st.date_input(
                    "Reprogramar para",
                    value=ref_date() + timedelta(days=1),
                    key=f"new_day_{prefix}_{activity_id}",
                )
                if st.button("Reprogramar", key=f"reschedule_{prefix}_{activity_id}"):
                    reschedule_activity(
                        activity_id,
                        user,
                        new_day,
                        note or "Atividade reprogramada",
                    )
                    st.rerun()

            with c:
                if st.button("Cancelar", key=f"cancel_{prefix}_{activity_id}"):
                    cancel_activity(
                        activity_id,
                        user,
                        note or "Atividade cancelada",
                    )
                    st.rerun()

def activity_groups(frame_data: pd.DataFrame, user: str, prefix: str) -> None:
    if frame_data.empty:
        st.success("Nenhuma atividade pendente nesta data.")
        return

    order = ["Em andamento", "Pendente", "Concluída", "Reprogramada", "Cancelada"]
    for status in order:
        section = frame_data[frame_data["status"] == status]
        if section.empty:
            continue
        st.markdown(f"### {status}")
        for _, row in section.iterrows():
            activity_card(row, user, f"{prefix}_{plain(status)}")


def home(user: str) -> None:
    day = frame("Meu Painel", f"Atividades de {user}")
    ensure_activities(day)

    mine = list_user_activities(day, user)
    previous = list_previous_pending(day, user)

    if mine.empty:
        pending = progress = completed = mine
    else:
        pending = mine[
            ~mine["status"].isin(["Concluída", "Cancelada", "Reprogramada"])
        ]
        progress = mine[mine["status"] == "Em andamento"]
        completed = mine[mine["status"] == "Concluída"]

    st.markdown(
        f"""
        <div class='hero'>
            <h1>Olá, {user} 👋</h1>
            <p>Suas atividades e prioridades de {br(day)}.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1: metric("Pendentes", len(pending), br(day), "#f59e0b")
    with c2: metric("Em andamento", len(progress), "Atividades iniciadas", "#f97316")
    with c3: metric("Pendências anteriores", len(previous), "Somente suas", "#dc2626")
    with c4: metric("Concluídas", len(completed), "Na data selecionada", "#16a34a")

    st.markdown("<div class='panel'><div class='panel-title'>⭐ Minhas prioridades</div>", unsafe_allow_html=True)
    activity_groups(pending, user, "home")
    st.markdown("</div>", unsafe_allow_html=True)

    if not previous.empty:
        st.markdown("<div class='panel'><div class='panel-title'>🔴 Minhas pendências anteriores</div>", unsafe_allow_html=True)
        activity_groups(previous, user, "home_previous")
        st.markdown("</div>", unsafe_allow_html=True)

def my_day(user: str) -> None:
    day = frame("Meu Dia", f"Atividades atribuídas a {user}")
    activities = list_user_activities(day, user)

    pending = activities[
        ~activities["status"].isin(["Concluída", "Cancelada", "Reprogramada"])
    ] if not activities.empty else activities
    progress = activities[activities["status"] == "Em andamento"] if not activities.empty else activities
    completed = activities[activities["status"] == "Concluída"] if not activities.empty else activities

    c1, c2, c3 = st.columns(3)
    with c1: metric("Pendentes", len(pending), br(day), "#f59e0b")
    with c2: metric("Em andamento", len(progress), "Atividades iniciadas", "#f97316")
    with c3: metric("Concluídas", len(completed), "Na data selecionada", "#16a34a")

    show_completed = st.toggle("Exibir atividades finalizadas", value=False)
    display = activities if show_completed else pending
    activity_groups(display, user, "myday")


def team(user: str) -> None:
    day = frame("Equipe", "Andamento por pessoa")
    activities = list_activities(day)
    users = list_users(True)

    if users.empty:
        st.info("Nenhum usuário cadastrado.")
        return

    summary_rows = []

    for person in users["name"].dropna().astype(str).tolist():
        if activities.empty:
            person_activities = activities.copy()
        else:
            person_activities = activities[
                activities["owners"].apply(
                    lambda value: owner_matches(value, person)
                )
            ].copy()

        total = len(person_activities)
        completed = len(
            person_activities[person_activities["status"] == "Concluída"]
        ) if total else 0
        in_progress = len(
            person_activities[person_activities["status"] == "Em andamento"]
        ) if total else 0
        pending = len(
            person_activities[
                ~person_activities["status"].isin(
                    ["Concluída", "Cancelada", "Reprogramada"]
                )
            ]
        ) if total else 0
        progress_pct = round(completed / total * 100) if total else 0
        previous = list_previous_pending(day, person)

        summary_rows.append(
            {
                "person": person,
                "department": user_profile(person).get("department", ""),
                "total": total,
                "pending": pending,
                "in_progress": in_progress,
                "completed": completed,
                "previous": len(previous),
                "progress_pct": progress_pct,
            }
        )

    total_programmed = sum(item["total"] for item in summary_rows)
    total_pending = sum(item["pending"] for item in summary_rows)
    total_completed = sum(item["completed"] for item in summary_rows)
    total_previous = sum(item["previous"] for item in summary_rows)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric("Programadas", total_programmed, br(day), "#2563eb")
    with c2:
        metric("Pendentes", total_pending, "Equipe", "#f59e0b")
    with c3:
        metric("Pendências anteriores", total_previous, "Equipe", "#dc2626")
    with c4:
        metric("Concluídas", total_completed, "Equipe", "#16a34a")

    st.markdown(
        "<div class='panel'><div class='panel-title'>👥 Colaboradores</div>",
        unsafe_allow_html=True,
    )

    for item in summary_rows:
        with st.container(border=True):
            left, center, right = st.columns([2.4, 2.8, 1.2])

            with left:
                st.markdown(f"### 👤 {item['person']}")
                st.caption(item["department"] or "Sem departamento")

            with center:
                p1, p2, p3, p4 = st.columns(4)
                with p1:
                    st.metric("Programadas", item["total"])
                with p2:
                    st.metric("Pendentes", item["pending"])
                with p3:
                    st.metric("Em andamento", item["in_progress"])
                with p4:
                    st.metric("Concluídas", item["completed"])

                st.progress(
                    min(max(item["progress_pct"], 0), 100) / 100,
                    text=f"{item['progress_pct']}% concluído",
                )

                if item["previous"] > 0:
                    st.warning(
                        f"{item['previous']} pendência(s) anterior(es)"
                    )

            with right:
                if st.button(
                    "Ver atividades",
                    key=f"team_view_{plain(item['person'])}",
                    use_container_width=True,
                ):
                    st.session_state["team_selected_person"] = item["person"]

    st.markdown("</div>", unsafe_allow_html=True)

    selected_person = st.session_state.get("team_selected_person")

    if selected_person:
        st.markdown(
            f"<div class='panel'><div class='panel-title'>📋 Atividades de {selected_person}</div>",
            unsafe_allow_html=True,
        )

        if activities.empty:
            selected = activities.copy()
        else:
            selected = activities[
                activities["owners"].apply(
                    lambda value: owner_matches(value, selected_person)
                )
            ].copy()

        show_finished = st.toggle(
            "Exibir atividades finalizadas",
            value=False,
            key="team_show_finished",
        )

        if not show_finished and not selected.empty:
            selected = selected[
                ~selected["status"].isin(
                    ["Concluída", "Cancelada", "Reprogramada"]
                )
            ]

        activity_groups(
            selected,
            user,
            f"team_{plain(selected_person)}",
        )

        if st.button("Fechar atividades", key="team_close_person"):
            st.session_state.pop("team_selected_person", None)
            st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

def coordination(user: str) -> None:
    day = frame("Coordenação", "Pendências e entregas da equipe")
    activities = list_activities(day)
    previous = list_previous_pending(day)

    open_activities = activities[
        ~activities["status"].isin(["Concluída", "Cancelada", "Reprogramada"])
    ] if not activities.empty else activities
    progress = activities[activities["status"] == "Em andamento"] if not activities.empty else activities
    completed = activities[activities["status"] == "Concluída"] if not activities.empty else activities

    c1, c2, c3, c4 = st.columns(4)
    with c1: metric("Abertas", len(open_activities), "Atividades do dia", "#f59e0b")
    with c2: metric("Em andamento", len(progress), "Atividades iniciadas", "#f97316")
    with c3: metric("Pendências anteriores", len(previous), "Requer atenção", "#dc2626")
    with c4: metric("Concluídas", len(completed), br(day), "#16a34a")

    st.markdown("<div class='panel'><div class='panel-title'>🔴 Pendências anteriores</div>", unsafe_allow_html=True)
    activity_groups(previous, user, "coord_previous")
    st.markdown("</div>", unsafe_allow_html=True)


def department_page(user: str, department: str) -> None:
    day = frame(department, f"Atividades de {department}")
    activities = list_department_activities(day, department)

    show_completed = st.toggle("Exibir atividades finalizadas", value=False)
    if not show_completed and not activities.empty:
        activities = activities[
            ~activities["status"].isin(["Concluída", "Cancelada", "Reprogramada"])
        ]

    activity_groups(activities, user, f"department_{plain(department)}")


def pending_page(user: str) -> None:
    day = frame("Pendências", "Atividades abertas de dias úteis anteriores")
    previous = list_previous_pending(day)
    activity_groups(previous, user, "previous")


def routine_form_values(routine: dict | None = None) -> dict:
    routine = routine or {}
    return {
        "title": normalize_text(routine.get("title")),
        "description": normalize_text(routine.get("description")),
        "department": normalize_text(routine.get("department")),
        "owners": [
            owner.strip()
            for owner in normalize_text(routine.get("owners"))
            .replace("/", ",")
            .replace(";", ",")
            .split(",")
            if owner.strip()
        ],
        "frequency": normalize_text(routine.get("frequency")) or "Diária",
        "due_rule": normalize_text(routine.get("due_rule")),
        "priority": normalize_text(routine.get("priority")) or "Normal",
        "mandatory": bool(routine.get("mandatory", 0)),
        "project": normalize_text(routine.get("project")),
        "start_date": pd.to_datetime(
            routine.get("start_date"),
            dayfirst=True,
            errors="coerce",
        ).date() if normalize_text(routine.get("start_date")) else ref_date(),
        "active": bool(routine.get("active", 1)),
    }


def render_routine_form(
    user: str,
    users: list[str],
    departments: list[str],
    projects: list[str],
    routine: dict | None = None,
    form_key: str = "routine_form",
    duplicate_mode: bool = False,
) -> None:
    values = routine_form_values(routine)
    editing = routine is not None and not duplicate_mode

    with st.form(form_key):
        a, b = st.columns(2)

        with a:
            title = st.text_input(
                "Atividade",
                value=(
                    f"{values['title']} - Cópia"
                    if duplicate_mode
                    else values["title"]
                ),
            )
            description = st.text_area(
                "Descrição",
                value=values["description"],
            )

            default_department = (
                departments.index(values["department"])
                if values["department"] in departments
                else 0
            )
            department = st.selectbox(
                "Departamento",
                departments,
                index=default_department,
            )

            selected_owners = [
                owner for owner in values["owners"] if owner in users
            ]
            owners = st.multiselect(
                "Responsáveis",
                users,
                default=selected_owners,
            )

            default_project = (
                projects.index(values["project"])
                if values["project"] in projects
                else 0
            )
            project = st.selectbox(
                "Projeto",
                projects,
                index=default_project,
            )

        with b:
            frequencies = [
                "Diária", "Semanal", "Quinzenal",
                "Mensal", "Anual", "Única",
            ]
            default_frequency = (
                frequencies.index(values["frequency"])
                if values["frequency"] in frequencies
                else 0
            )
            frequency = st.selectbox(
                "Periodicidade",
                frequencies,
                index=default_frequency,
            )

            help_text = {
                "Diária": "Horário opcional, por exemplo: 17:30",
                "Semanal": "Informe o dia da semana ou use a data inicial como referência.",
                "Quinzenal": "Ex.: Dia 05 / Dia 18",
                "Mensal": "Ex.: Dia 10",
                "Anual": "Ex.: 15/03",
                "Única": "A execução ocorrerá somente na data de início.",
            }.get(frequency, "")

            due_rule = st.text_input(
                "Regra de execução",
                value=values["due_rule"],
                help=help_text,
                placeholder="Ex.: Dia 10, Dia 05 / Dia 18 ou 17:30",
            )

            priorities = ["Normal", "Alta", "Crítica", "Baixa"]
            default_priority = (
                priorities.index(values["priority"])
                if values["priority"] in priorities
                else 0
            )
            priority = st.selectbox(
                "Prioridade",
                priorities,
                index=default_priority,
            )

            mandatory = st.checkbox(
                "Obrigatória",
                value=values["mandatory"],
            )
            start_date = st.date_input(
                "Data de início",
                value=values["start_date"],
            )

            active = st.checkbox(
                "Ativa",
                value=values["active"],
                disabled=not editing,
            )

        label = (
            "Salvar alterações"
            if editing
            else "Criar rotina"
        )
        submitted = st.form_submit_button(
            label,
            type="primary",
            use_container_width=True,
        )

    if submitted:
        if not title.strip():
            st.error("Informe o nome da atividade.")
            return
        if not owners:
            st.error("Selecione pelo menos um responsável.")
            return

        payload = {
            "source_id": (
                routine.get("source_id")
                if editing
                else f"manual-{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
            ),
            "title": title.strip(),
            "description": description.strip(),
            "department": department,
            "owners": "/".join(owners),
            "frequency": frequency,
            "due_rule": due_rule.strip(),
            "priority": priority,
            "mandatory": mandatory,
            "project": project,
            "start_date": br(start_date),
            "active": active if editing else 1,
        }

        if editing:
            update_routine(int(routine["id"]), payload, user)
            st.session_state.pop("edit_routine_id", None)
            st.success("Rotina atualizada.")
        else:
            create_snapshot()
            create_routine(payload, user)
            st.success("Rotina cadastrada.")

        st.rerun()


def routines_page(user: str) -> None:
    day = frame("Rotinas", "Cadastro das atividades recorrentes")

    users = list_users(True)["name"].dropna().astype(str).tolist()
    departments = list_departments() or [
        "Contas a Receber",
        "Contas a Pagar",
        "Contabilidade",
        "Controladoria",
        "Tesouraria",
    ]
    projects_frame = list_projects(True)
    projects = [""] + (
        projects_frame["name"].dropna().astype(str).tolist()
        if not projects_frame.empty
        else []
    )

    st.markdown("### Manutenção de rotinas")

    if st.button(
        "➕ Nova rotina",
        type="primary",
        use_container_width=False,
    ):
        st.session_state["show_new_routine"] = True
        st.session_state.pop("edit_routine_id", None)

    st.caption(
        "Cadastre, edite, duplique, desative ou reative atividades recorrentes."
    )

    if st.session_state.get("show_new_routine"):
        with st.expander("➕ Nova rotina", expanded=True):
            render_routine_form(
                user,
                users,
                departments,
                projects,
                routine=None,
                form_key="new_routine_form",
            )
            if st.button("Fechar cadastro", key="close_new_routine"):
                st.session_state.pop("show_new_routine", None)
                st.rerun()

    edit_id = st.session_state.get("edit_routine_id")
    if edit_id:
        routine = get_routine(int(edit_id))
        if routine:
            with st.expander(
                f"✏️ Editar: {routine['title']}",
                expanded=True,
            ):
                render_routine_form(
                    user,
                    users,
                    departments,
                    projects,
                    routine=routine,
                    form_key=f"edit_routine_form_{edit_id}",
                )
                if st.button(
                    "Cancelar edição",
                    key=f"cancel_edit_{edit_id}",
                ):
                    st.session_state.pop("edit_routine_id", None)
                    st.rerun()

    routines = list_routines(False)

    if routines.empty:
        st.info("Nenhuma rotina cadastrada.")
        return

    st.markdown("### 🔎 Localizar rotinas")
    f1, f2, f3, f4 = st.columns(4)

    with f1:
        search = st.text_input(
            "Pesquisar",
            placeholder="Atividade, responsável ou projeto...",
        )
    with f2:
        department_filter = st.selectbox(
            "Departamento",
            ["Todos"] + departments,
        )
    with f3:
        owner_filter = st.selectbox(
            "Responsável",
            ["Todos"] + users,
        )
    with f4:
        status_filter = st.selectbox(
            "Situação",
            ["Ativas", "Inativas", "Todas"],
        )

    f5, f6 = st.columns(2)
    with f5:
        frequency_filter = st.selectbox(
            "Periodicidade",
            ["Todas", "Diária", "Semanal", "Quinzenal", "Mensal", "Anual", "Única"],
        )
    with f6:
        project_filter = st.selectbox(
            "Projeto",
            ["Todos"] + [p for p in projects if p],
        )

    filtered = routines.copy()

    if search.strip():
        needle = search.strip().casefold()
        filtered = filtered[
            filtered.apply(
                lambda row: needle in " ".join(
                    [
                        normalize_text(row.get("title")),
                        normalize_text(row.get("description")),
                        normalize_text(row.get("owners")),
                        normalize_text(row.get("project")),
                    ]
                ).casefold(),
                axis=1,
            )
        ]

    if department_filter != "Todos":
        filtered = filtered[
            filtered["department"].astype(str) == department_filter
        ]

    if owner_filter != "Todos":
        filtered = filtered[
            filtered["owners"].apply(
                lambda value: owner_matches(value, owner_filter)
            )
        ]

    if status_filter == "Ativas":
        filtered = filtered[filtered["active"] == 1]
    elif status_filter == "Inativas":
        filtered = filtered[filtered["active"] == 0]

    if frequency_filter != "Todas":
        filtered = filtered[
            filtered["frequency"].astype(str) == frequency_filter
        ]

    if project_filter != "Todos":
        filtered = filtered[
            filtered["project"].astype(str) == project_filter
        ]

    st.caption(
        f"{len(filtered)} rotina(s) encontrada(s)."
    )

    for _, row in filtered.iterrows():
        routine_id = int(row["id"])
        active = bool(row.get("active", 1))

        with st.container(border=True):
            left, actions = st.columns([4.8, 1.7])

            with left:
                status_label = "Ativa" if active else "Inativa"
                status_class_name = "green" if active else "gray"

                st.markdown(
                    f"### 📋 {normalize_text(row.get('title'))}"
                )

                if normalize_text(row.get("description")):
                    st.caption(normalize_text(row.get("description")))

                tags = (
                    f"<span class='tag {status_class_name}'>{status_label}</span>"
                    f"<span class='tag blue'>🏢 {normalize_text(row.get('department'))}</span>"
                    f"<span class='tag green'>👤 {normalize_text(row.get('owners'))}</span>"
                    f"<span class='tag purple'>🔁 {normalize_text(row.get('frequency'))}</span>"
                )

                if normalize_text(row.get("due_rule")):
                    tags += (
                        f"<span class='tag yellow'>📅 "
                        f"{normalize_text(row.get('due_rule'))}</span>"
                    )

                if normalize_text(row.get("project")):
                    tags += (
                        f"<span class='tag purple'>📁 "
                        f"{normalize_text(row.get('project'))}</span>"
                    )

                st.markdown(tags, unsafe_allow_html=True)

            with actions:
                if st.button(
                    "✏️ Editar",
                    key=f"edit_routine_{routine_id}",
                    use_container_width=True,
                ):
                    st.session_state["edit_routine_id"] = routine_id
                    st.session_state.pop("show_new_routine", None)
                    st.rerun()

                if st.button(
                    "📄 Duplicar",
                    key=f"duplicate_routine_{routine_id}",
                    use_container_width=True,
                ):
                    new_id = duplicate_routine(routine_id, user)
                    st.session_state["edit_routine_id"] = new_id
                    st.success("Cópia criada. Ajuste os campos e salve.")
                    st.rerun()

                toggle_label = (
                    "⏸️ Desativar"
                    if active
                    else "▶️ Reativar"
                )
                if st.button(
                    toggle_label,
                    key=f"toggle_routine_{routine_id}",
                    use_container_width=True,
                ):
                    set_routine_active(
                        routine_id,
                        not active,
                        user,
                    )
                    st.rerun()

                if not active:
                    with st.expander("🗑️ Excluir definitivamente"):
                        confirm = st.checkbox(
                            "Confirmo a exclusão desta rotina",
                            key=f"confirm_delete_{routine_id}",
                        )
                        if st.button(
                            "Excluir",
                            key=f"delete_routine_{routine_id}",
                            disabled=not confirm,
                            use_container_width=True,
                        ):
                            deleted = delete_routine_permanently(
                                routine_id,
                                user,
                            )
                            if deleted:
                                st.success("Rotina excluída.")
                                st.rerun()
                            else:
                                st.error(
                                    "Esta rotina já possui registros diários "
                                    "e não pode ser excluída. Mantenha-a desativada."
                                )

def projects_page(user: str) -> None:
    day = frame("Projetos", "Prazos e andamento dos projetos")
    projects = list_projects(True)

    if projects.empty:
        st.info("Nenhum projeto cadastrado.")
        return

    st.dataframe(
        projects[["name", "department", "owners", "status", "progress", "due_date", "next_step"]],
        use_container_width=True,
        hide_index=True,
    )


def history_page(user: str) -> None:
    day = frame("Histórico", "Ações registradas no sistema")
    show_all = st.toggle("Exibir todo o histórico", value=False)
    events = list_events(None if show_all else day)
    st.dataframe(events, use_container_width=True, hide_index=True)


def administration_page(user: str) -> None:
    day = frame("Usuários", "Cadastro e atualização da equipe")

    users = list_users(False)
    c1, c2, c3 = st.columns(3)
    with c1: metric("Usuários", len(users), "Cadastrados", "#2563eb")
    with c2: metric("Rotinas", len(list_routines(False)), "Cadastradas", "#7c3aed")
    with c3: metric("Projetos", len(list_projects(False)), "Cadastrados", "#f59e0b")

    st.info(
        f"📅 Início oficial da operação: **{br(GO_LIVE_DATE)}**. "
        "Indicadores e pendências anteriores consideram somente dias úteis "
        "a partir desta data."
    )

    with st.expander("➕ Cadastrar usuário"):
        with st.form("new_user"):
            name = st.text_input("Nome")
            role = st.selectbox("Perfil", ["Usuário", "Coordenador", "Administradora"], help="O perfil fica registrado para organização, mas não limita o acesso nesta versão.")
            department = st.selectbox(
                "Departamento",
                list_departments() or ["Controladoria"],
            )
            save = st.form_submit_button("Salvar usuário")
        if save and name:
            create_user(name, role, department)
            st.success("Usuário cadastrado.")
            st.rerun()

    st.dataframe(users[["name", "role", "department", "active"]], use_container_width=True, hide_index=True)


def backup_page(user: str) -> None:
    day = frame("Backup", "Cópias, exportação e importação da base")

    export_path = Path("FIRST_OPS_Backup.xlsx")

    st.markdown("### Criar e baixar cópias")

    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button(
            "Criar backup completo",
            type="primary",
            use_container_width=True,
        ):
            create_full_backup_package()
            cleanup_old_backups(30)
            st.success("Backup completo criado.")

    with c2:
        if st.button(
            "Gerar arquivo Excel",
            use_container_width=True,
        ):
            export_excel(export_path)
            st.success("Arquivo Excel gerado.")

    with c3:
        if st.button(
            "Criar cópia do banco",
            use_container_width=True,
        ):
            path = create_snapshot()
            cleanup_old_backups(30)
            st.success(f"Cópia criada: {path.name}")

    if export_path.exists():
        st.download_button(
            "Baixar backup Excel",
            data=export_path.read_bytes(),
            file_name=f"FIRST_OPS_Backup_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if DB_PATH.exists():
        st.download_button(
            "Baixar banco SQLite",
            data=DB_PATH.read_bytes(),
            file_name=f"first_ops_{date.today().isoformat()}.db",
            mime="application/octet-stream",
            use_container_width=True,
        )

    st.markdown("### Importar rotinas")

    st.caption(
        "O arquivo deve conter as colunas: title, owners, description, "
        "department, frequency, due_rule, priority, mandatory, project e start_date."
    )

    upload = st.file_uploader(
        "Selecione o Excel de rotinas",
        type=["xlsx"],
        key="backup_routine_import",
    )

    if upload is not None:
        imported = pd.read_excel(upload, dtype=object)
        required = {"title", "owners"}
        missing = required.difference(imported.columns)

        if missing:
            st.error(
                "Colunas obrigatórias ausentes: "
                + ", ".join(sorted(missing))
            )
        else:
            preview_columns = [
                column for column in [
                    "title", "description", "department", "owners",
                    "frequency", "due_rule", "priority", "project"
                ] if column in imported.columns
            ]

            st.dataframe(
                imported[preview_columns].head(20),
                use_container_width=True,
                hide_index=True,
            )

            if st.button(
                "Importar rotinas",
                type="primary",
                use_container_width=True,
            ):
                create_full_backup_package()
                imported_count = 0
                errors = []

                for index, row in imported.iterrows():
                    try:
                        title = normalize_text(row.get("title"))
                        owners_value = normalize_text(row.get("owners"))

                        if not title or not owners_value:
                            continue

                        create_routine(
                            {
                                "source_id": (
                                    normalize_text(row.get("source_id"))
                                    or f"excel-{datetime.now().strftime('%Y%m%d%H%M%S%f')}-{index}"
                                ),
                                "title": title,
                                "description": normalize_text(row.get("description")),
                                "department": normalize_text(row.get("department")),
                                "owners": owners_value,
                                "frequency": normalize_text(row.get("frequency")) or "Diária",
                                "due_rule": normalize_text(row.get("due_rule")),
                                "priority": normalize_text(row.get("priority")) or "Normal",
                                "mandatory": str(row.get("mandatory", "")).strip().lower() in {
                                    "1", "sim", "s", "true"
                                },
                                "project": normalize_text(row.get("project")),
                                "start_date": normalize_text(row.get("start_date")) or br(day),
                            },
                            user,
                        )
                        imported_count += 1
                    except Exception as exc:
                        errors.append(f"Linha {index + 2}: {exc}")

                st.success(
                    f"{imported_count} rotina(s) importada(s)."
                )

                if errors:
                    st.warning(
                        "Linhas não importadas:\n"
                        + "\n".join(errors[:10])
                    )

                st.rerun()

    st.markdown("### Cópias disponíveis")

    backups = list_backup_files()

    if backups.empty:
        st.caption("Nenhuma cópia local criada.")
    else:
        st.dataframe(
            backups[["arquivo", "tamanho_kb", "modificado_em"]],
            use_container_width=True,
            hide_index=True,
        )

    st.caption(
        "O sistema mantém as 30 cópias mais recentes e cria uma cópia "
        "antes de alterações importantes."
    )

    st.success(
        f"Início oficial dos indicadores: {br(GO_LIVE_DATE)}."
    )

def main() -> None:
    user, page, departments = sidebar()

    if page == "Meu Painel":
        home(user)
    elif page == "Meu Dia":
        my_day(user)
    elif page == "Equipe":
        team(user)
    elif page == "Coordenação":
        coordination(user)
    elif page in departments:
        department_page(user, page)
    elif page == "Pendências":
        pending_page(user)
    elif page == "Rotinas":
        routines_page(user)
    elif page == "Projetos":
        projects_page(user)
    elif page == "Histórico":
        history_page(user)
    elif page == "Usuários":
        administration_page(user)
    elif page == "Backup":
        backup_page(user)


if __name__ == "__main__":
    main()
