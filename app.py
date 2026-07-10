
# -*- coding: utf-8 -*-
from datetime import date, timedelta
from pathlib import Path
import unicodedata

import pandas as pd
import streamlit as st

from db import (
    DB_PATH, init_db, query_df, setting, set_setting,
    list_users, create_user, list_departments,
    list_routines, create_routine, archive_routine,
    list_executions, list_previous_pending, create_execution_from_routine,
    set_execution_status, add_execution_note, reschedule_execution, cancel_execution,
    list_projects, create_project, list_events, export_excel
)
from migrate import preview_backup, migrate_backup, repair_from_backup


st.set_page_config(
    page_title="FIRST OPS 3.2",
    page_icon="⚕️",
    layout="wide",
    initial_sidebar_state="expanded"
)

init_db(DB_PATH)

st.markdown("""
<style>
.stApp { background:#f4f7fb; }
div.block-container { padding-top:1.4rem; padding-bottom:4rem; max-width:1550px; }
section[data-testid="stSidebar"] { background:linear-gradient(180deg,#020617,#111827 55%,#312e81); }
section[data-testid="stSidebar"] * { color:#fff !important; }
section[data-testid="stSidebar"] div[data-baseweb="select"] * { color:#0f172a !important; }
.ops-title { font-size:42px; font-weight:950; color:#0f172a; letter-spacing:-.04em; }
.ops-sub { color:#64748b; margin-bottom:18px; }
.hero { background:linear-gradient(135deg,#0f172a,#312e81 55%,#7c3aed); color:#fff; border-radius:32px; padding:30px 34px; box-shadow:0 22px 65px rgba(49,46,129,.25); margin-bottom:22px; }
.hero h1 { margin:0; font-size:34px; font-weight:950; }
.hero p { color:#ddd6fe; margin:8px 0 0 0; }
.panel { background:#fff; border:1px solid #e2e8f0; border-radius:26px; padding:22px; box-shadow:0 14px 36px rgba(15,23,42,.06); margin-top:18px; }
.panel-title { font-size:22px; font-weight:900; color:#0f172a; margin-bottom:15px; }
.metric { background:#fff; border:1px solid #e2e8f0; border-radius:24px; padding:22px; min-height:125px; box-shadow:0 14px 34px rgba(15,23,42,.06); }
.metric-label { color:#64748b; text-transform:uppercase; font-size:12px; font-weight:850; }
.metric-value { font-size:38px; font-weight:950; margin-top:8px; }
.tag { display:inline-block; padding:6px 10px; border-radius:999px; font-size:12px; font-weight:850; margin-right:6px; margin-top:6px; }
.blue{background:#dbeafe;color:#1d4ed8}.green{background:#dcfce7;color:#15803d}.red{background:#fee2e2;color:#b91c1c}.yellow{background:#fef3c7;color:#b45309}.purple{background:#ede9fe;color:#6d28d9}.gray{background:#f1f5f9;color:#475569}.orange{background:#ffedd5;color:#c2410c}
.task-title{font-size:18px;font-weight:900;color:#0f172a}.done{text-decoration:line-through;color:#94a3b8}
.stButton>button{border-radius:14px;font-weight:850}

section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea,
section[data-testid="stSidebar"] select,
section[data-testid="stSidebar"] [data-baseweb="select"] span,
section[data-testid="stSidebar"] [data-baseweb="select"] div {
    color:#0f172a !important;
}

section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background:#ffffff !important;
    border:1px solid #cbd5e1 !important;
    border-radius:12px !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] span {
    color:#0f172a !important;
    font-weight:700 !important;
}

</style>
""", unsafe_allow_html=True)


def text(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def norm(v):
    out = text(v).lower()
    out = unicodedata.normalize("NFKD", out).encode("ascii", "ignore").decode("ascii")
    return " ".join(out.split())


def _menu_norm(value):
    return norm(value)



def user_profile(user):
    users = list_users(True)
    if users.empty:
        return {"role": "Usuário", "department": ""}

    selected = users[users["name"].astype(str).str.lower() == str(user).lower()]
    if selected.empty:
        return {"role": "Usuário", "department": ""}

    row = selected.iloc[0]
    return {
        "role": text(row.get("role")) or "Usuário",
        "department": text(row.get("department")),
    }


def is_admin(user):
    role = norm(user_profile(user).get("role"))
    return role in ["administrador", "administradora", "admin"]


def owner_matches(owner_value, user):
    owner = text(owner_value)
    target = norm(user)

    if not owner or not target:
        return False

    names = owner.replace("/", ",").replace(";", ",").replace("|", ",").split(",")
    normalized_names = [norm(name) for name in names if text(name)]

    return target in normalized_names


def filter_user_tasks(dataframe, user):
    if dataframe.empty or "owner" not in dataframe.columns:
        return dataframe

    return dataframe[dataframe["owner"].apply(lambda value: owner_matches(value, user))]


def today():
    return date.today()


def br(d):
    return d.strftime("%d/%m/%Y")


def parse_br(v):
    if not text(v):
        return None
    try:
        return pd.to_datetime(v, dayfirst=True).date()
    except Exception:
        return None


def is_business_day(d):
    return d.weekday() < 5


def ref_date():
    if "ref_date" not in st.session_state:
        st.session_state["ref_date"] = today()
    return st.session_state["ref_date"]


def move_date(days):
    st.session_state["ref_date"] = ref_date() + timedelta(days=days)


def go_today():
    st.session_state["ref_date"] = today()


def frequency(value):
    x = norm(value)
    if x in ["", "diaria", "diario", "daily", "todo dia", "todos os dias"]:
        return "diaria"
    if x in ["semanal", "semana", "weekly"]:
        return "semanal"
    if x in ["mensal", "mes", "monthly"]:
        return "mensal"
    if x in ["unica", "unico", "pontual"]:
        return "unica"
    return x


def should_generate(routine, d):
    if not is_business_day(d):
        return False

    start = parse_br(routine.get("start_date"))
    if start and start > d:
        return False

    f = frequency(routine.get("frequency"))
    if f == "diaria":
        return True
    if f == "semanal":
        base = start or d
        return (d - base).days % 7 == 0
    if f == "mensal":
        base = start or d
        return d.day == base.day
    if f == "unica":
        return start == d
    return True


def generate_day(d):
    routines = list_routines(True)
    if routines.empty:
        return 0
    count = 0
    for _, row in routines.iterrows():
        if should_generate(row, d):
            create_execution_from_routine(row, br(d))
            count += 1
    return count


def metric(label, value, detail, color):
    st.markdown(f"""
    <div class='metric'>
        <div class='metric-label'>{label}</div>
        <div class='metric-value' style='color:{color}'>{value}</div>
        <div style='color:{color};font-size:13px'>{detail}</div>
    </div>
    """, unsafe_allow_html=True)


def header(title, subtitle):
    st.markdown(f"<div class='ops-title'>{title}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='ops-sub'>{subtitle}</div>", unsafe_allow_html=True)


def date_bar():
    st.markdown("<div class='panel' style='padding:14px 18px'>", unsafe_allow_html=True)
    c1,c2,c3,c4,c5 = st.columns([1.1,1.2,.85,.7,2.5])
    with c1:
        st.markdown("#### 📅 Data de referência")
    with c2:
        st.session_state["ref_date"] = st.date_input("Data", value=ref_date(), label_visibility="collapsed")
    with c3:
        st.button("◀ Anterior", use_container_width=True, on_click=move_date, args=(-1,))
    with c4:
        st.button("Hoje", use_container_width=True, on_click=go_today)
    with c5:
        cc1,cc2 = st.columns([.8,1.7])
        with cc1:
            st.button("Próximo ▶", use_container_width=True, on_click=move_date, args=(1,))
        with cc2:
            if is_business_day(ref_date()):
                st.success(f"Dia útil: {br(ref_date())}")
            else:
                st.warning("Data sem rotina operacional automática.")
    st.markdown("</div>", unsafe_allow_html=True)
    return ref_date()


def frame(title, subtitle):
    header(title, subtitle)
    return date_bar()


def ensure_admin_seed():
    if list_users(True).empty:
        create_user("Paula", "Administradora", "Coordenação")


def sidebar():
    ensure_admin_seed()

    st.sidebar.markdown("## ⚕️ FIRST OPS")
    st.sidebar.markdown("### Gestão Operacional")
    st.sidebar.divider()

    users = list_users(True)
    names = users["name"].dropna().astype(str).tolist() if not users.empty else []

    if not names:
        create_user("Paula", "Administradora", "Coordenação")
        names = ["Paula"]

    user = st.sidebar.selectbox("Usuário", names)
    profile = user_profile(user)

    st.sidebar.caption(
        f"{profile.get('role', 'Usuário')} • "
        f"{profile.get('department') or 'Sem departamento'}"
    )

    fixed = {
        "inicio", "meu dia", "equipe", "coordenacao", "pendencias",
        "rotinas", "projetos", "historico", "administracao",
        "base e backup", "exportar"
    }

    deps = list_departments()
    deps = sorted(
        list(
            dict.fromkeys(
                [d for d in deps if text(d) and _menu_norm(d) not in fixed]
            )
        )
    )

    operational_menu = ["Início", "Meu Dia", "Equipe", "Coordenação"]
    operational_menu += deps
    operational_menu += ["Pendências", "Rotinas", "Projetos", "Histórico"]

    if is_admin(user):
        operational_menu += ["Administração", "Base e Backup", "Exportar"]

    page = st.sidebar.radio("Navegação", operational_menu)

    return user, page, deps

def status_class(status):
    return {
        "Pendente":"yellow",
        "Em andamento":"orange",
        "Concluída":"green",
        "Cancelada":"gray",
        "Reprogramada":"blue"
    }.get(status, "purple")


def execution_card(row, user, prefix):
    status = text(row.get("status")) or "Pendente"
    eid = int(row.get("id"))
    with st.container(border=True):
        c1,c2 = st.columns([4.8,1.4])
        with c1:
            cls = "task-title done" if status == "Concluída" else "task-title"
            st.markdown(f"<div class='{cls}'>{text(row.get('title'))}</div>", unsafe_allow_html=True)
            if text(row.get("description")):
                st.caption(text(row.get("description")))

            tags = f"<span class='tag {status_class(status)}'>{status}</span>"
            if text(row.get("department")):
                tags += f"<span class='tag blue'>{text(row.get('department'))}</span>"
            if text(row.get("owner")):
                tags += f"<span class='tag green'>👤 {text(row.get('owner'))}</span>"
            if text(row.get("project")):
                tags += f"<span class='tag purple'>📁 {text(row.get('project'))}</span>"
            if text(row.get("priority")):
                tags += f"<span class='tag yellow'>{text(row.get('priority'))}</span>"
            st.markdown(tags, unsafe_allow_html=True)

            if text(row.get("note")):
                st.info(text(row.get("note")))

        with c2:
            if status == "Concluída":
                st.success("Concluída")
            else:
                if status != "Em andamento":
                    if st.button("▶ Iniciar", key=f"start_{prefix}_{eid}"):
                        set_execution_status(eid, "Em andamento", user, "Iniciada")
                        st.rerun()
                if st.button("✅ Concluir", key=f"done_{prefix}_{eid}"):
                    set_execution_status(eid, "Concluída", user, "Concluída")
                    st.rerun()

        with st.expander("💬 Justificar / reprogramar / cancelar"):
            note = st.text_area("Observação", key=f"note_{prefix}_{eid}")
            cA,cB,cC = st.columns(3)
            with cA:
                if st.button("Salvar observação", key=f"save_note_{prefix}_{eid}"):
                    add_execution_note(eid, user, note)
                    st.rerun()
            with cB:
                new_date = st.date_input("Nova data", value=ref_date()+timedelta(days=1), key=f"resch_date_{prefix}_{eid}")
                if st.button("Reprogramar", key=f"resch_{prefix}_{eid}"):
                    reschedule_execution(eid, user, br(new_date), note or "Reprogramada")
                    st.rerun()
            with cC:
                if st.button("Cancelar", key=f"cancel_{prefix}_{eid}"):
                    cancel_execution(eid, user, note or "Cancelada")
                    st.rerun()


def get_day_data(d):
    day = list_executions(br(d))
    previous = list_previous_pending(br(d))
    if day.empty:
        return day, day, day, day, previous

    open_day = day[~day["status"].isin(["Concluída","Cancelada","Reprogramada"])]
    done_day = day[day["status"] == "Concluída"]
    progress = day[day["status"] == "Em andamento"]
    return day, open_day, done_day, progress, previous


def home(user):
    d = frame("Visão Geral", "Acompanhamento das atividades da equipe")

    # A criação usa INSERT OR IGNORE, portanto pode ser chamada sem duplicar tarefas.
    if is_business_day(d):
        generate_day(d)

    st.markdown(f"""
    <div class='hero'>
        <h1>Olá, {user} 👋</h1>
        <p>Atividades e prioridades de {br(d)}.</p>
    </div>
    """, unsafe_allow_html=True)

    day, open_day, done_day, progress, previous = get_day_data(d)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        metric("Programadas", len(day), br(d), "#2563eb")
    with c2:
        metric("Pendentes", len(open_day), "Aguardando conclusão", "#f59e0b")
    with c3:
        metric("Em andamento", len(progress), "Atividades iniciadas", "#f97316")
    with c4:
        metric("Pendências anteriores", len(previous), "Requer atenção", "#dc2626")
    with c5:
        metric("Concluídas", len(done_day), "Na data selecionada", "#16a34a")

    st.markdown(
        "<div class='panel'><div class='panel-title'>⭐ Minhas prioridades</div>",
        unsafe_allow_html=True
    )

    mine = filter_user_tasks(open_day, user)

    if mine.empty:
        st.success("Você não possui atividades pendentes nesta data.")
    else:
        for _, row in mine.iterrows():
            execution_card(row, user, "home_mine")

    st.markdown("</div>", unsafe_allow_html=True)

def migracao(user):
    header("Base e Backup", "Importação inicial e manutenção da base")

    preview = preview_backup("Agenda.xlsx")

    if not preview.get("found"):
        st.error("Não encontrei o arquivo Agenda.xlsx na pasta do app.")
        return

    st.markdown("<div class='panel'><div class='panel-title'>📦 Arquivo de origem</div>", unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    with c1: metric("Usuários", preview.get("users", 0), "Backup", "#2563eb")
    with c2: metric("Rotinas", preview.get("tasks", 0), "Virarão rotinas mestre", "#7c3aed")
    with c3: metric("Projetos", preview.get("projects", 0), "Backup", "#f59e0b")
    with c4: metric("Histórico", preview.get("history", 0), "Auditoria", "#16a34a")
    st.caption("Abas encontradas: " + ", ".join(preview.get("sheets", [])))
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='panel'><div class='panel-title'>🔎 Dados disponíveis no sistema</div>", unsafe_allow_html=True)
    cA, cB, cC = st.columns(3)
    with cA:
        metric("Usuários no banco", len(list_users(False)), "FIRST OPS", "#2563eb")
    with cB:
        metric("Rotinas no banco", len(list_routines(False)), "FIRST OPS", "#7c3aed")
    with cC:
        metric("Projetos no banco", len(list_projects(False)), "FIRST OPS", "#f59e0b")
    st.markdown("</div>", unsafe_allow_html=True)

    done = setting("migration_done", "0") == "1"
    if done:
        st.success("A importação inicial já foi realizada.")
    else:
        st.warning("A importação traz usuários, rotinas, projetos e registros históricos.")

    st.markdown("<div class='panel'><div class='panel-title'>🛠️ Reprocessar arquivo de origem</div>", unsafe_allow_html=True)
    st.info("Use esta opção somente para complementar cadastros que não tenham sido importados.")
    if st.button("REPROCESSAR CADASTROS"):
        ok, msg = repair_from_backup("Agenda.xlsx")
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.warning(msg)
    st.markdown("</div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("IMPORTAR ARQUIVO", type="primary", disabled=done):
            ok, msg = migrate_backup("Agenda.xlsx", force=False)
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.warning(msg)

    with col2:
        with st.expander("Importação administrativa"):
            st.warning("Use apenas se estiver testando e souber que o banco pode receber nova carga.")
            senha = st.text_input("Senha", type="password")
            if st.button("Forçar migração"):
                if senha == "Paula2026":
                    ok, msg = migrate_backup("Agenda.xlsx", force=True)
                    st.success(msg if ok else msg)
                    st.rerun()
                else:
                    st.error("Senha incorreta.")


def meu_dia(user):
    d = frame("Meu Dia", f"Atividades atribuídas a {user}")

    if is_business_day(d):
        generate_day(d)

    day, open_day, done_day, progress, previous = get_day_data(d)
    mine = filter_user_tasks(day, user)

    c1, c2, c3 = st.columns(3)

    if mine.empty:
        mine_open = mine
        mine_progress = mine
        mine_done = mine
    else:
        mine_open = mine[
            ~mine["status"].isin(["Concluída", "Cancelada", "Reprogramada"])
        ]
        mine_progress = mine[mine["status"] == "Em andamento"]
        mine_done = mine[mine["status"] == "Concluída"]

    with c1:
        metric("Pendentes", len(mine_open), br(d), "#f59e0b")
    with c2:
        metric("Em andamento", len(mine_progress), "Atividades iniciadas", "#f97316")
    with c3:
        metric("Concluídas", len(mine_done), "Na data selecionada", "#16a34a")

    show_done = st.toggle("Exibir atividades concluídas", value=False)

    display = mine.copy()
    if not show_done and not display.empty:
        display = display[
            ~display["status"].isin(["Concluída", "Cancelada", "Reprogramada"])
        ]

    if display.empty:
        st.success("Nenhuma atividade pendente para você nesta data.")
        return

    for _, row in display.iterrows():
        execution_card(row, user, "meu_dia")

def equipe(user):
    d = frame("Equipe", "Distribuição e andamento das atividades")

    if is_business_day(d):
        generate_day(d)

    day, open_day, done_day, progress, previous = get_day_data(d)

    if day.empty:
        st.info("Sem checklist gerado para esta data.")
        return

    rows = []
    for owner in sorted([x for x in day["owner"].dropna().unique() if text(x)]):
        base = day[day["owner"] == owner]
        total = len(base)
        done = len(base[base["status"] == "Concluída"])
        open_count = len(base[~base["status"].isin(["Concluída","Cancelada","Reprogramada"])])
        pct = round((done/total)*100, 1) if total else 0
        rows.append({"Usuário": owner, "Total": total, "Abertas": open_count, "Concluídas": done, "Progresso %": pct})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def coordenacao(user):
    d = frame("Coordenação", "Acompanhamento das entregas da equipe")

    if is_business_day(d):
        generate_day(d)

    day, open_day, done_day, progress, previous = get_day_data(d)

    c1,c2,c3,c4 = st.columns(4)
    with c1: metric("Abertas", len(open_day), "Executar", "#f59e0b")
    with c2: metric("Em andamento", len(progress), "Iniciadas", "#f97316")
    with c3: metric("Pendências", len(previous), "Dias anteriores", "#dc2626")
    with c4: metric("Concluídas", len(done_day), br(d), "#16a34a")

    st.markdown("<div class='panel'><div class='panel-title'>🔴 Pendências anteriores</div>", unsafe_allow_html=True)
    if previous.empty:
        st.success("Sem pendências anteriores.")
    else:
        for _, row in previous.iterrows():
            execution_card(row, user, "coord_prev")
    st.markdown("</div>", unsafe_allow_html=True)


def department_page(user, department):
    d = frame(department, f"Atividades de {department}")

    if is_business_day(d):
        generate_day(d)

    day = list_executions(br(d))
    base = day[day["department"].astype(str).str.lower() == department.lower()] if not day.empty else day
    if base.empty:
        st.info("Sem tarefas para este departamento nesta data.")
        return
    show_done = st.toggle("Mostrar concluídas", value=False)
    if not show_done:
        base = base[~base["status"].isin(["Concluída","Cancelada","Reprogramada"])]
    for _, row in base.iterrows():
        execution_card(row, user, f"dep_{department}")


def pendencias(user):
    d = frame("Pendências", "Atividades de datas anteriores ainda abertas")
    previous = list_previous_pending(br(d))
    if previous.empty:
        st.success("Sem pendências anteriores.")
    else:
        for _, row in previous.iterrows():
            execution_card(row, user, "pend")


def rotinas(user):
    d = frame("Rotinas", "Atividades recorrentes da operação")
    with st.expander("➕ Nova rotina", expanded=False):
        users = list_users()
        owners = [""] + (users["name"].tolist() if not users.empty else [])
        deps = list_departments() or ["Contas a Receber", "Contas a Pagar", "Contabilidade", "Controladoria", "Tesouraria"]
        projects = list_projects()
        projects_list = [""] + (projects["name"].tolist() if not projects.empty else [])
        with st.form("new_routine"):
            c1,c2 = st.columns(2)
            with c1:
                title = st.text_input("Rotina")
                desc = st.text_area("Descrição")
                dep = st.selectbox("Departamento", deps)
                owner = st.selectbox("Responsável", owners)
                project = st.selectbox("Projeto", projects_list)
            with c2:
                frequency = st.selectbox("Periodicidade", ["Diária", "Semanal", "Mensal", "Única"])
                priority = st.selectbox("Prioridade", ["Normal", "Alta", "Crítica", "Baixa"])
                mandatory = st.selectbox("Obrigatória", ["Não", "Sim"])
                start = st.date_input("Data de início", value=d)
            save = st.form_submit_button("Salvar rotina")
        if save:
            create_routine({
                "title": title, "description": desc, "department": dep, "owner": owner,
                "frequency": frequency, "priority": priority, "mandatory": 1 if mandatory=="Sim" else 0,
                "start_date": br(start), "project": project
            }, user)
            st.success("Rotina criada.")
            st.rerun()

    routines = list_routines(True)
    st.dataframe(routines, use_container_width=True, hide_index=True)


def projetos(user):
    d = frame("Projetos", "Acompanhamento dos projetos da área")
    with st.expander("➕ Novo projeto", expanded=False):
        users = list_users()
        owners = [""] + (users["name"].tolist() if not users.empty else [])
        with st.form("new_project"):
            name = st.text_input("Projeto")
            desc = st.text_area("Descrição")
            dep = st.text_input("Departamento")
            owner = st.selectbox("Responsável", owners)
            due = st.date_input("Prazo final", value=d)
            stage = st.selectbox("Etapa", ["Planejamento", "Em andamento", "Validação", "Concluído", "Suspenso"])
            next_step = st.text_input("Próxima etapa")
            save = st.form_submit_button("Salvar projeto")
        if save:
            create_project({
                "name": name, "description": desc, "department": dep, "owner": owner,
                "due_date": br(due), "stage": stage, "next_step": next_step
            }, user)
            st.success("Projeto criado.")
            st.rerun()

    projects = list_projects(True)
    if projects.empty:
        st.info("Nenhum projeto cadastrado.")
        return

    cols = st.columns(4)
    stages = ["Planejamento", "Em andamento", "Validação", "Concluído"]
    for i, stage in enumerate(stages):
        with cols[i]:
            st.markdown(f"### {stage}")
            subset = projects[projects["stage"] == stage]
            if subset.empty:
                st.caption("Sem projetos.")
            else:
                for _, row in subset.iterrows():
                    with st.container(border=True):
                        st.markdown(f"**📁 {row['name']}**")
                        st.caption(f"{row.get('owner','')} | Prazo: {row.get('due_date','')}")
                        if text(row.get("next_step")):
                            st.info(row.get("next_step"))


def historico():
    d = frame("Histórico", "Registro das movimentações realizadas")
    events = list_events()
    if events.empty:
        st.info("Sem histórico.")
        return
    show_all = st.toggle("Mostrar tudo", value=False)
    if not show_all:
        events = events[events["event_date"].astype(str) == br(d)]
    st.dataframe(events, use_container_width=True, hide_index=True)


def administracao(user):
    d = frame("Administração", "Usuários e configurações de acesso")
    c1,c2,c3 = st.columns(3)
    with c1: st.metric("Usuários", len(list_users(False)))
    with c2: st.metric("Rotinas", len(list_routines(False)))
    with c3: st.metric("Execuções", len(list_executions(include_all=True)))

    st.markdown("<div class='panel'><div class='panel-title'>👥 Usuários</div>", unsafe_allow_html=True)
    with st.form("new_user"):
        name = st.text_input("Nome")
        role = st.selectbox("Perfil", ["Usuário", "Administradora"])
        dep = st.text_input("Departamento")
        save = st.form_submit_button("Criar usuário")
    if save and name:
        create_user(name, role, dep)
        st.success("Usuário criado.")
        st.rerun()
    st.dataframe(list_users(False), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='panel'><div class='panel-title'>⚙️ Status da Migração</div>", unsafe_allow_html=True)
    st.write(f"Migração realizada: **{setting('migration_done','0')}**")
    st.write(f"Arquivo: **{setting('migration_file','-')}**")
    st.markdown("</div>", unsafe_allow_html=True)


def exportar(user):
    d = frame("Exportar", "Gerar arquivo para conferência e segurança")
    path = Path("FIRST_OPS_Export.xlsx")
    if st.button("Gerar Excel"):
        export_excel(path)
        st.success("Excel gerado.")
    if path.exists():
        st.download_button("Baixar Excel", data=path.read_bytes(), file_name=path.name)


def main():
    user, page, deps = sidebar()

    if page == "Início":
        home(user)
    elif page == "Meu Dia":
        meu_dia(user)
    elif page == "Equipe":
        equipe(user)
    elif page == "Coordenação":
        coordenacao(user)
    elif page in deps:
        department_page(user, page)
    elif page == "Pendências":
        pendencias(user)
    elif page == "Rotinas":
        rotinas(user)
    elif page == "Projetos":
        projetos(user)
    elif page == "Histórico":
        historico()
    elif page == "Administração" and is_admin(user):
        administracao(user)
    elif page == "Base e Backup" and is_admin(user):
        migracao(user)
    elif page == "Exportar" and is_admin(user):
        exportar(user)
    else:
        st.error("Acesso não autorizado para este perfil.")


if __name__ == "__main__":
    main()
