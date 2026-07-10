
# -*- coding: utf-8 -*-
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path
import unicodedata

import pandas as pd
import streamlit as st

from first_ops_database import (
    BACKUP_DIR,
    DB_PATH,
    backfill_until,
    cancel_activity,
    create_project,
    create_routine,
    create_snapshot,
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
    user_profile,
)
from migrate import import_backup


st.set_page_config(
    page_title="FIRST OPS Enterprise 2.0",
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

    menu = ["Início", "Meu Dia", "Equipe", "Coordenação"]
    menu += departments
    menu += ["Pendências", "Rotinas", "Projetos", "Histórico"]

    if is_admin(user):
        menu += ["Administração", "Base e Backup"]

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
    title_class = "task-done" if status == "Concluída" else "task-title"

    with st.container(border=True):
        left, right = st.columns([4.8, 1.35])

        with left:
            description = normalize_text(row.get("description"))
            title = normalize_text(row.get("title"))

            st.markdown(
                f"<div class='{title_class}'>{title}</div>",
                unsafe_allow_html=True,
            )

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
            else:
                if status != "Em andamento":
                    if st.button("▶ Iniciar", key=f"start_{prefix}_{activity_id}"):
                        set_activity_status(activity_id, "Em andamento", user, "Atividade iniciada")
                        st.rerun()

                if st.button("✅ Concluir", key=f"done_{prefix}_{activity_id}"):
                    set_activity_status(activity_id, "Concluída", user, "Atividade concluída")
                    st.rerun()

        with st.expander("💬 Observação, reprogramação ou cancelamento"):
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
                    reschedule_activity(activity_id, user, new_day, note or "Atividade reprogramada")
                    st.rerun()

            with c:
                if st.button("Cancelar", key=f"cancel_{prefix}_{activity_id}"):
                    cancel_activity(activity_id, user, note or "Atividade cancelada")
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
    day = frame("Visão Geral", "Acompanhamento das atividades da equipe")
    ensure_activities(day)

    activities = list_activities(day)
    pending = activities[
        ~activities["status"].isin(["Concluída", "Cancelada", "Reprogramada"])
    ] if not activities.empty else activities
    progress = activities[activities["status"] == "Em andamento"] if not activities.empty else activities
    completed = activities[activities["status"] == "Concluída"] if not activities.empty else activities
    previous = list_previous_pending(day)

    st.markdown(
        f"""
        <div class='hero'>
            <h1>Olá, {user} 👋</h1>
            <p>Atividades e prioridades de {br(day)}.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: metric("Programadas", len(activities), br(day), "#2563eb")
    with c2: metric("Pendentes", len(pending), "Aguardando conclusão", "#f59e0b")
    with c3: metric("Em andamento", len(progress), "Atividades iniciadas", "#f97316")
    with c4: metric("Pendências anteriores", len(previous), "Requer atenção", "#dc2626")
    with c5: metric("Concluídas", len(completed), "Na data selecionada", "#16a34a")

    st.markdown("<div class='panel'><div class='panel-title'>⭐ Minhas prioridades</div>", unsafe_allow_html=True)
    mine = list_user_activities(day, user)
    mine = mine[~mine["status"].isin(["Concluída", "Cancelada", "Reprogramada"])] if not mine.empty else mine
    activity_groups(mine, user, "home")
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
    day = frame("Equipe", "Distribuição e andamento das atividades")
    activities = list_activities(day)

    if activities.empty:
        st.info("Nenhuma atividade programada nesta data.")
        return

    rows = []
    for person in list_users(True)["name"].tolist():
        person_activities = activities[
            activities["owners"].apply(lambda value: owner_matches(value, person))
        ]
        if person_activities.empty:
            continue

        total = len(person_activities)
        done = len(person_activities[person_activities["status"] == "Concluída"])
        open_count = len(
            person_activities[
                ~person_activities["status"].isin(["Concluída", "Cancelada", "Reprogramada"])
            ]
        )
        rows.append(
            {
                "Pessoa": person,
                "Departamento": user_profile(person).get("department", ""),
                "Programadas": total,
                "Pendentes": open_count,
                "Concluídas": done,
                "Progresso %": round(done / total * 100, 1) if total else 0,
            }
        )

    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def coordination(user: str) -> None:
    day = frame("Coordenação", "Acompanhamento das entregas da equipe")
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
    day = frame("Pendências", "Atividades de datas anteriores ainda abertas")
    previous = list_previous_pending(day)
    activity_groups(previous, user, "previous")


def routines_page(user: str) -> None:
    day = frame("Rotinas", "Atividades recorrentes da operação")
    users = list_users(True)["name"].tolist()
    departments = list_departments() or [
        "Contas a Receber", "Contas a Pagar", "Contabilidade",
        "Controladoria", "Tesouraria"
    ]
    project_names = [""] + (
        list_projects(True)["name"].tolist() if not list_projects(True).empty else []
    )

    if is_admin(user):
        with st.expander("➕ Cadastrar rotina", expanded=False):
            with st.form("new_routine"):
                a, b = st.columns(2)
                with a:
                    title = st.text_input("Atividade")
                    description = st.text_area("Descrição")
                    department = st.selectbox("Departamento", departments)
                    owners = st.multiselect("Responsáveis", users)
                    project = st.selectbox("Projeto", project_names)
                with b:
                    frequency = st.selectbox("Periodicidade", ["Diária", "Mensal", "Quinzenal", "Semanal", "Única"])
                    due_rule = st.text_input("Dia ou horário", placeholder="Ex.: Dia 10, Dia 05 / Dia 18 ou 17:30")
                    priority = st.selectbox("Prioridade", ["Normal", "Alta", "Crítica", "Baixa"])
                    mandatory = st.checkbox("Obrigatória")
                    start_date = st.date_input("Data de início", value=day)
                save = st.form_submit_button("Salvar rotina")

            if save:
                if not title or not owners:
                    st.error("Informe a atividade e pelo menos um responsável.")
                else:
                    create_routine(
                        {
                            "title": title,
                            "description": description,
                            "department": department,
                            "owners": "/".join(owners),
                            "frequency": frequency,
                            "due_rule": due_rule,
                            "priority": priority,
                            "mandatory": mandatory,
                            "project": project,
                            "start_date": br(start_date),
                        },
                        user,
                    )
                    st.success("Rotina cadastrada.")
                    st.rerun()

    routines = list_routines(True)
    columns = ["id", "title", "description", "department", "owners", "frequency", "due_rule", "priority", "project"]
    st.dataframe(routines[columns], use_container_width=True, hide_index=True)


def projects_page(user: str) -> None:
    day = frame("Projetos", "Acompanhamento dos projetos da área")
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
    day = frame("Histórico", "Registro diário das movimentações")
    show_all = st.toggle("Exibir todo o histórico", value=False)
    events = list_events(None if show_all else day)
    st.dataframe(events, use_container_width=True, hide_index=True)


def administration_page(user: str) -> None:
    day = frame("Administração", "Usuários e cadastros")
    if not is_admin(user):
        st.error("Acesso restrito.")
        return

    users = list_users(False)
    c1, c2, c3 = st.columns(3)
    with c1: metric("Usuários", len(users), "Cadastrados", "#2563eb")
    with c2: metric("Rotinas", len(list_routines(False)), "Cadastradas", "#7c3aed")
    with c3: metric("Projetos", len(list_projects(False)), "Cadastrados", "#f59e0b")

    with st.expander("➕ Cadastrar usuário"):
        with st.form("new_user"):
            name = st.text_input("Nome")
            role = st.selectbox("Perfil", ["Usuário", "Coordenador", "Administradora"])
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
    day = frame("Base e Backup", "Segurança e exportação dos registros")
    if not is_admin(user):
        st.error("Acesso restrito.")
        return

    export_path = Path("FIRST_OPS_Backup.xlsx")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Gerar backup Excel", type="primary"):
            export_excel(export_path)
            st.success("Backup Excel gerado.")
        if export_path.exists():
            st.download_button(
                "Baixar backup Excel",
                data=export_path.read_bytes(),
                file_name=f"FIRST_OPS_Backup_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with c2:
        if st.button("Criar cópia do banco"):
            path = create_snapshot()
            st.success(f"Cópia criada: {path.name}")
        if DB_PATH.exists():
            st.download_button(
                "Baixar banco SQLite",
                data=DB_PATH.read_bytes(),
                file_name=f"first_ops_{date.today().isoformat()}.db",
                mime="application/octet-stream",
            )

    st.info(
        "O sistema cria automaticamente uma cópia local do banco por dia. "
        "Use os botões acima para baixar uma cópia externa."
    )


def main() -> None:
    user, page, departments = sidebar()

    if page == "Início":
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
    elif page == "Administração":
        administration_page(user)
    elif page == "Base e Backup":
        backup_page(user)


if __name__ == "__main__":
    main()
