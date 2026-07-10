# -*- coding: utf-8 -*-
from datetime import date, timedelta
from pathlib import Path
import unicodedata

import pandas as pd
import streamlit as st

from first_ops_database import (
    DB_PATH, init_db, setting,
    list_users, create_user, list_departments,
    list_routines, create_routine, archive_routine,
    list_executions, create_execution_from_routine, materialize_execution,
    set_execution_status, add_execution_note, reschedule_execution, cancel_execution,
    list_projects, create_project, list_events, export_excel,
)
from migrate import preview_backup, migrate_backup, repair_from_backup

st.set_page_config(
    page_title="FIRST OPS Enterprise 1.2",
    page_icon="⚕️",
    layout="wide",
    initial_sidebar_state="expanded",
)
init_db(DB_PATH)

st.markdown("""
<style>
.stApp { background:#f4f7fb; }
div.block-container { padding-top:1.25rem; padding-bottom:4rem; max-width:1550px; }
section[data-testid="stSidebar"] { background:linear-gradient(180deg,#020617,#111827 55%,#312e81); }
section[data-testid="stSidebar"] * { color:#fff !important; }
section[data-testid="stSidebar"] div[data-baseweb="select"] > div { background:#fff !important; border-radius:12px !important; }
section[data-testid="stSidebar"] div[data-baseweb="select"] span { color:#0f172a !important; font-weight:700 !important; }
.ops-title { font-size:40px; font-weight:950; color:#0f172a; letter-spacing:-.04em; }
.ops-sub { color:#64748b; margin-bottom:18px; }
.hero { background:linear-gradient(135deg,#0f172a,#312e81 55%,#7c3aed); color:#fff; border-radius:30px; padding:28px 32px; box-shadow:0 22px 60px rgba(49,46,129,.22); margin-bottom:22px; }
.hero h1 { margin:0; font-size:32px; font-weight:950; }
.hero p { color:#ddd6fe; margin:8px 0 0; }
.panel { background:#fff; border:1px solid #e2e8f0; border-radius:24px; padding:20px; box-shadow:0 14px 34px rgba(15,23,42,.06); margin-top:18px; }
.panel-title { font-size:21px; font-weight:900; color:#0f172a; margin-bottom:14px; }
.metric { background:#fff; border:1px solid #e2e8f0; border-radius:22px; padding:20px; min-height:120px; box-shadow:0 12px 30px rgba(15,23,42,.06); }
.metric-label { color:#64748b; text-transform:uppercase; font-size:12px; font-weight:850; }
.metric-value { font-size:36px; font-weight:950; margin-top:8px; }
.tag { display:inline-block; padding:6px 10px; border-radius:999px; font-size:12px; font-weight:850; margin-right:6px; margin-top:6px; }
.blue{background:#dbeafe;color:#1d4ed8}.green{background:#dcfce7;color:#15803d}.red{background:#fee2e2;color:#b91c1c}.yellow{background:#fef3c7;color:#b45309}.purple{background:#ede9fe;color:#6d28d9}.gray{background:#f1f5f9;color:#475569}.orange{background:#ffedd5;color:#c2410c}
.task-title{font-size:18px;font-weight:900;color:#0f172a}.done{text-decoration:line-through;color:#94a3b8}
.stButton>button{border-radius:13px;font-weight:800}

section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 12px !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] input,
section[data-testid="stSidebar"] div[data-baseweb="select"] span,
section[data-testid="stSidebar"] div[data-baseweb="select"] div {
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    font-weight: 700 !important;
}
section[data-testid="stSidebar"] div[data-baseweb="select"] svg {
    fill: #0f172a !important;
    color: #0f172a !important;
}

</style>
""", unsafe_allow_html=True)


def text(v):
    if v is None or pd.isna(v):
        return ""
    return str(v).strip()


def norm(v):
    out = text(v).lower()
    out = unicodedata.normalize("NFKD", out).encode("ascii", "ignore").decode("ascii")
    return " ".join(out.split())


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


def previous_business_day(d):
    d = d - timedelta(days=1)
    while not is_business_day(d):
        d -= timedelta(days=1)
    return d


def frequency(value):
    value = norm(value)
    if value in {"", "diaria", "diario", "todo dia", "todos os dias", "daily"}:
        return "diaria"
    if value in {"semanal", "semana", "weekly"}:
        return "semanal"
    if value in {"mensal", "mes", "monthly"}:
        return "mensal"
    if value in {"unica", "unico", "pontual"}:
        return "unica"
    return value


def should_show(routine, d):
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


def ref_date():
    if "ref_date" not in st.session_state:
        st.session_state.ref_date = today()
    return st.session_state.ref_date


def set_date_widget(d):
    st.session_state.ref_date = d
    st.session_state.date_selector = d


def move_date(days):
    set_date_widget(ref_date() + timedelta(days=days))


def go_today():
    set_date_widget(today())


def header(title, subtitle):
    st.markdown(f"<div class='ops-title'>{title}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='ops-sub'>{subtitle}</div>", unsafe_allow_html=True)


def metric(label, value, detail, color):
    st.markdown(f"""
    <div class='metric'>
      <div class='metric-label'>{label}</div>
      <div class='metric-value' style='color:{color}'>{value}</div>
      <div style='color:{color};font-size:13px'>{detail}</div>
    </div>
    """, unsafe_allow_html=True)


def date_bar():
    if "date_selector" not in st.session_state:
        st.session_state.date_selector = ref_date()
    st.markdown("<div class='panel' style='padding:14px 18px'>", unsafe_allow_html=True)
    c1,c2,c3,c4,c5 = st.columns([1.1,1.2,.85,.7,2.5])
    with c1:
        st.markdown("#### 📅 Data de referência")
    with c2:
        selected = st.date_input("Data", key="date_selector", label_visibility="collapsed")
        st.session_state.ref_date = selected
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


def ensure_admin():
    users = list_users(True)
    if users.empty:
        create_user("Paula", "Administradora", "Controladoria")


def user_profile(user):
    users = list_users(True)
    selected = users[users["name"].astype(str).str.lower() == str(user).lower()] if not users.empty else users
    if selected.empty:
        return {"role":"Usuário", "department":""}
    row = selected.iloc[0]
    return {"role":text(row.get("role")) or "Usuário", "department":text(row.get("department"))}


def is_admin(user):
    return norm(user_profile(user)["role"]) in {"administrador", "administradora", "admin"}


def owner_matches(value, user):
    parts = text(value).replace("/", ",").replace(";", ",").replace("|", ",").split(",")
    return norm(user) in {norm(x) for x in parts if text(x)}


def filter_user(df, user):
    if df.empty:
        return df
    return df[df["owner"].apply(lambda x: owner_matches(x, user))]


def sidebar():
    ensure_admin()
    st.sidebar.markdown("## ⚕️ FIRST OPS")
    st.sidebar.markdown("### Gestão Operacional • 1.2")
    st.sidebar.divider()
    users = list_users(True)
    names = users["name"].dropna().astype(str).tolist() if not users.empty else ["Paula"]
    user = st.sidebar.selectbox("Usuário", names)
    profile = user_profile(user)
    st.sidebar.caption(f"{profile['role']} • {profile['department'] or 'Sem departamento'}")

    reserved = {"inicio","meu dia","equipe","coordenacao","pendencias","rotinas","projetos","historico","administracao","base e backup","exportar"}
    deps = sorted(dict.fromkeys([d for d in list_departments() if text(d) and norm(d) not in reserved]))
    menu = ["Início","Meu Dia","Equipe","Coordenação"] + deps + ["Pendências","Rotinas","Projetos","Histórico"]
    if is_admin(user):
        menu += ["Administração","Base e Backup","Exportar"]
    page = st.sidebar.radio("Navegação", menu)
    return user, page, deps


def virtual_day(d):
    """Combina rotinas previstas e execuções reais sem gravar ao abrir a tela."""
    routines = list_routines(True)
    actual = list_executions(br(d))
    actual_by_routine = {}
    if not actual.empty:
        for _, ex in actual.iterrows():
            rid = ex.get("routine_id")
            if pd.notna(rid):
                actual_by_routine[int(rid)] = ex.to_dict()

    rows = []
    due_ids = set()
    if not routines.empty:
        for _, routine in routines.iterrows():
            if not should_show(routine, d):
                continue
            rid = int(routine["id"])
            due_ids.add(rid)
            if rid in actual_by_routine:
                row = actual_by_routine[rid]
                row["virtual"] = False
            else:
                row = {
                    "id": -rid,
                    "routine_id": rid,
                    "execution_date": br(d),
                    "title": routine.get("title", ""),
                    "description": routine.get("description", ""),
                    "department": routine.get("department", ""),
                    "owner": routine.get("owner", ""),
                    "project": routine.get("project", ""),
                    "priority": routine.get("priority", "Normal"),
                    "mandatory": routine.get("mandatory", 0),
                    "status": "Pendente",
                    "note": "",
                    "virtual": True,
                }
            rows.append(row)

    # Execuções reprogramadas/manuais que não pertencem à recorrência calculada também aparecem.
    if not actual.empty:
        for _, ex in actual.iterrows():
            rid = ex.get("routine_id")
            if pd.isna(rid) or int(rid) not in due_ids:
                row = ex.to_dict()
                row["virtual"] = False
                rows.append(row)

    columns = ["id","routine_id","execution_date","title","description","department","owner","project","priority","mandatory","status","note","virtual"]
    return pd.DataFrame(rows, columns=columns) if rows else pd.DataFrame(columns=columns)


def previous_pending(d):
    """Pendências do último dia útil, incluindo rotinas sem interação."""
    prior = previous_business_day(d)
    df = virtual_day(prior)
    if df.empty:
        return df
    return df[~df["status"].isin(["Concluída","Cancelada","Reprogramada"])].copy()


def real_execution_id(row):
    if bool(row.get("virtual", False)) or int(row.get("id", 0)) < 0:
        return materialize_execution(int(row.get("routine_id")), text(row.get("execution_date")))
    return int(row.get("id"))


def status_class(status):
    return {"Pendente":"yellow","Em andamento":"orange","Concluída":"green","Cancelada":"gray","Reprogramada":"blue"}.get(status,"purple")


def execution_card(row, user, prefix):
    status = text(row.get("status")) or "Pendente"
    key_id = f"r{int(row.get('routine_id') or 0)}_{text(row.get('execution_date')).replace('/','')}"
    with st.container(border=True):
        c1,c2 = st.columns([4.8,1.4])
        with c1:
            cls = "task-title done" if status == "Concluída" else "task-title"
            st.markdown(f"<div class='{cls}'>{text(row.get('title'))}</div>", unsafe_allow_html=True)
            if text(row.get("description")):
                st.caption(text(row.get("description")))
            tags = f"<span class='tag {status_class(status)}'>{status}</span>"
            if text(row.get("department")): tags += f"<span class='tag blue'>{text(row.get('department'))}</span>"
            if text(row.get("owner")): tags += f"<span class='tag green'>👤 {text(row.get('owner'))}</span>"
            if text(row.get("project")): tags += f"<span class='tag purple'>📁 {text(row.get('project'))}</span>"
            if text(row.get("priority")): tags += f"<span class='tag yellow'>{text(row.get('priority'))}</span>"
            st.markdown(tags, unsafe_allow_html=True)
            if bool(row.get("virtual", False)):
                st.caption("Programada pela rotina • será registrada quando houver uma ação")
            if text(row.get("note")):
                st.info(text(row.get("note")))
        with c2:
            if status == "Concluída":
                st.success("Concluída")
            else:
                if status != "Em andamento" and st.button("▶ Iniciar", key=f"start_{prefix}_{key_id}"):
                    eid = real_execution_id(row)
                    set_execution_status(eid, "Em andamento", user, "Iniciada")
                    st.rerun()
                if st.button("✅ Concluir", key=f"done_{prefix}_{key_id}"):
                    eid = real_execution_id(row)
                    set_execution_status(eid, "Concluída", user, "Concluída")
                    st.rerun()
        with st.expander("💬 Observação, reprogramação ou cancelamento"):
            note = st.text_area("Observação", key=f"note_{prefix}_{key_id}")
            cA,cB,cC = st.columns(3)
            with cA:
                if st.button("Salvar observação", key=f"save_{prefix}_{key_id}"):
                    eid = real_execution_id(row)
                    add_execution_note(eid, user, note)
                    st.rerun()
            with cB:
                new_date = st.date_input("Nova data", value=ref_date()+timedelta(days=1), key=f"date_{prefix}_{key_id}")
                if st.button("Reprogramar", key=f"resch_{prefix}_{key_id}"):
                    eid = real_execution_id(row)
                    reschedule_execution(eid, user, br(new_date), note or "Reprogramada")
                    st.rerun()
            with cC:
                if st.button("Cancelar", key=f"cancel_{prefix}_{key_id}"):
                    eid = real_execution_id(row)
                    cancel_execution(eid, user, note or "Cancelada")
                    st.rerun()


def day_data(d):
    day = virtual_day(d)
    prev = previous_pending(d)
    opened = day[~day["status"].isin(["Concluída","Cancelada","Reprogramada"])] if not day.empty else day
    done = day[day["status"] == "Concluída"] if not day.empty else day
    progress = day[day["status"] == "Em andamento"] if not day.empty else day
    return day, opened, done, progress, prev


def home(user):
    d = frame("Visão Geral", "Atividades e prioridades da equipe")
    st.markdown(f"<div class='hero'><h1>Olá, {user} 👋</h1><p>Visão operacional de {br(d)}.</p></div>", unsafe_allow_html=True)
    day, opened, done, progress, prev = day_data(d)
    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: metric("Programadas", len(day), br(d), "#2563eb")
    with c2: metric("Pendentes", len(opened), "Aguardando conclusão", "#f59e0b")
    with c3: metric("Em andamento", len(progress), "Atividades iniciadas", "#f97316")
    with c4: metric("Pendências anteriores", len(prev), "Último dia útil", "#dc2626")
    with c5: metric("Concluídas", len(done), "Na data selecionada", "#16a34a")
    st.markdown("<div class='panel'><div class='panel-title'>⭐ Minhas prioridades</div>", unsafe_allow_html=True)
    mine = filter_user(opened, user)
    if mine.empty: st.success("Você não possui atividades pendentes nesta data.")
    else:
        for _, row in mine.iterrows(): execution_card(row, user, "home")
    st.markdown("</div>", unsafe_allow_html=True)


def meu_dia(user):
    d = frame("Meu Dia", f"Atividades atribuídas a {user}")
    day, _, _, _, _ = day_data(d)
    mine = filter_user(day, user)
    open_m = mine[~mine["status"].isin(["Concluída","Cancelada","Reprogramada"])] if not mine.empty else mine
    progress = mine[mine["status"] == "Em andamento"] if not mine.empty else mine
    done = mine[mine["status"] == "Concluída"] if not mine.empty else mine
    c1,c2,c3 = st.columns(3)
    with c1: metric("Pendentes", len(open_m), br(d), "#f59e0b")
    with c2: metric("Em andamento", len(progress), "Atividades iniciadas", "#f97316")
    with c3: metric("Concluídas", len(done), "Na data selecionada", "#16a34a")
    show_done = st.toggle("Exibir atividades concluídas", value=False)
    display = mine if show_done else open_m
    if display.empty: st.success("Nenhuma atividade pendente para você nesta data.")
    else:
        for _, row in display.iterrows(): execution_card(row, user, "mine")


def equipe(user):
    d = frame("Equipe", "Distribuição e andamento das atividades")
    day = virtual_day(d)
    if day.empty: st.info("Nenhuma atividade programada para esta data."); return
    rows=[]
    for owner in sorted([x for x in day["owner"].dropna().unique() if text(x)]):
        base=day[day["owner"]==owner]; total=len(base); complete=len(base[base["status"]=="Concluída"])
        rows.append({"Responsável":owner,"Programadas":total,"Pendentes":len(base[~base["status"].isin(["Concluída","Cancelada","Reprogramada"])]),"Concluídas":complete,"Progresso %":round(complete/total*100,1) if total else 0})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def coordenacao(user):
    d = frame("Coordenação", "Acompanhamento das entregas da equipe")
    day, opened, done, progress, prev = day_data(d)
    c1,c2,c3,c4 = st.columns(4)
    with c1: metric("Pendentes",len(opened),"Na data selecionada","#f59e0b")
    with c2: metric("Em andamento",len(progress),"Atividades iniciadas","#f97316")
    with c3: metric("Pendências anteriores",len(prev),"Último dia útil","#dc2626")
    with c4: metric("Concluídas",len(done),br(d),"#16a34a")
    st.markdown("<div class='panel'><div class='panel-title'>🔴 Pendências anteriores</div>", unsafe_allow_html=True)
    if prev.empty: st.success("Sem pendências do último dia útil.")
    else:
        for _, row in prev.iterrows(): execution_card(row,user,"coordprev")
    st.markdown("</div>", unsafe_allow_html=True)


def department_page(user, department):
    d = frame(department, f"Atividades de {department}")
    day=virtual_day(d); base=day[day["department"].astype(str).str.lower()==department.lower()] if not day.empty else day
    if base.empty: st.info("Nenhuma atividade programada para este departamento."); return
    show=st.toggle("Exibir concluídas",value=False)
    if not show: base=base[~base["status"].isin(["Concluída","Cancelada","Reprogramada"])]
    for _, row in base.iterrows(): execution_card(row,user,"dep")


def pendencias(user):
    d=frame("Pendências","Atividades do último dia útil ainda abertas")
    prev=previous_pending(d)
    if prev.empty: st.success("Sem pendências anteriores.")
    else:
        for _, row in prev.iterrows(): execution_card(row,user,"pend")


def rotinas(user):
    d=frame("Rotinas","Atividades recorrentes da operação")
    with st.expander("➕ Nova rotina",expanded=False):
        users=list_users(); owners=[""]+(users["name"].tolist() if not users.empty else []); deps=list_departments() or ["Contas a Receber","Contas a Pagar","Contabilidade","Controladoria","Tesouraria"]
        projects=list_projects(); project_options=[""]+(projects["name"].tolist() if not projects.empty else [])
        with st.form("new_routine"):
            c1,c2=st.columns(2)
            with c1:
                title=st.text_input("Atividade"); description=st.text_area("Descrição"); department=st.selectbox("Departamento",deps); owner=st.selectbox("Responsável",owners); project=st.selectbox("Projeto",project_options)
            with c2:
                freq=st.selectbox("Periodicidade",["Diária","Semanal","Mensal","Única"]); priority=st.selectbox("Prioridade",["Normal","Alta","Crítica","Baixa"]); mandatory=st.selectbox("Obrigatória",["Não","Sim"]); start=st.date_input("Data de início",value=d)
            save=st.form_submit_button("Salvar rotina")
        if save:
            create_routine({"title":title,"description":description,"department":department,"owner":owner,"frequency":freq,"priority":priority,"mandatory":1 if mandatory=="Sim" else 0,"start_date":br(start),"project":project},user); st.success("Rotina criada."); st.rerun()
    st.dataframe(list_routines(True),use_container_width=True,hide_index=True)


def projetos(user):
    d=frame("Projetos","Acompanhamento dos projetos da área")
    projects=list_projects(True)
    if projects.empty: st.info("Nenhum projeto cadastrado."); return
    stages=["Planejamento","Em andamento","Validação","Concluído"]; cols=st.columns(4)
    for i,stage in enumerate(stages):
        with cols[i]:
            st.markdown(f"### {stage}"); subset=projects[projects["stage"]==stage]
            if subset.empty: st.caption("Sem projetos.")
            else:
                for _,row in subset.iterrows():
                    with st.container(border=True): st.markdown(f"**📁 {row['name']}**"); st.caption(f"{row.get('owner','')} | Prazo: {row.get('due_date','')}"); st.info(row.get("next_step")) if text(row.get("next_step")) else None


def historico():
    d=frame("Histórico","Registro das movimentações realizadas"); events=list_events();
    if events.empty: st.info("Sem movimentações registradas."); return
    show=st.toggle("Mostrar todo o histórico",value=False)
    if not show: events=events[events["event_date"].astype(str)==br(d)]
    st.dataframe(events,use_container_width=True,hide_index=True)


def administracao(user):
    frame("Administração","Usuários e configurações de acesso")
    c1,c2,c3=st.columns(3)
    with c1: st.metric("Usuários",len(list_users(False)))
    with c2: st.metric("Rotinas",len(list_routines(False)))
    with c3: st.metric("Registros de execução",len(list_executions(include_all=True)))
    with st.form("new_user"):
        name=st.text_input("Nome"); role=st.selectbox("Perfil",["Usuário","Gestor","Coordenador","Administradora"]); dep=st.text_input("Departamento"); save=st.form_submit_button("Criar usuário")
    if save and name: create_user(name,role,dep); st.success("Usuário criado."); st.rerun()
    st.dataframe(list_users(False),use_container_width=True,hide_index=True)


def migracao(user):
    header("Base e Backup","Importação inicial e manutenção da base")
    preview=preview_backup("Agenda.xlsx")
    if not preview.get("found"): st.error("Arquivo Agenda.xlsx não encontrado."); return
    c1,c2,c3,c4=st.columns(4)
    with c1: metric("Usuários",preview.get("users",0),"Arquivo de origem","#2563eb")
    with c2: metric("Rotinas",preview.get("tasks",0),"Arquivo de origem","#7c3aed")
    with c3: metric("Projetos",preview.get("projects",0),"Arquivo de origem","#f59e0b")
    with c4: metric("Histórico",preview.get("history",0),"Arquivo de origem","#16a34a")
    if setting("migration_done","0")=="1": st.success("A importação inicial já foi realizada.")
    else:
        if st.button("IMPORTAR ARQUIVO",type="primary"):
            ok,msg=migrate_backup("Agenda.xlsx",force=False); st.success(msg) if ok else st.warning(msg); st.rerun()
    with st.expander("Complementar cadastros"):
        st.info("Reprocessa o arquivo sem apagar dados existentes.")
        if st.button("REPROCESSAR CADASTROS"):
            ok,msg=repair_from_backup("Agenda.xlsx"); st.success(msg) if ok else st.warning(msg); st.rerun()


def exportar(user):
    frame("Exportar","Gerar arquivo para conferência e segurança"); path=Path("FIRST_OPS_Export.xlsx")
    if st.button("Gerar arquivo"): export_excel(path); st.success("Arquivo gerado.")
    if path.exists(): st.download_button("Baixar arquivo",data=path.read_bytes(),file_name=path.name)


def main():
    user,page,deps=sidebar()
    if page=="Início": home(user)
    elif page=="Meu Dia": meu_dia(user)
    elif page=="Equipe": equipe(user)
    elif page=="Coordenação": coordenacao(user)
    elif page in deps: department_page(user,page)
    elif page=="Pendências": pendencias(user)
    elif page=="Rotinas": rotinas(user)
    elif page=="Projetos": projetos(user)
    elif page=="Histórico": historico()
    elif page=="Administração" and is_admin(user): administracao(user)
    elif page=="Base e Backup" and is_admin(user): migracao(user)
    elif page=="Exportar" and is_admin(user): exportar(user)
    else: st.error("Acesso não autorizado.")

if __name__=="__main__": main()
