
# -*- coding: utf-8 -*-
from pathlib import Path
from datetime import datetime, date, time
import pandas as pd
import streamlit as st

# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Agenda Operacional",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded"
)

ARQUIVO_EXCEL = "Agenda.xlsx"

ABAS = {
    "Usuarios": ["ID", "Nome", "Login", "Senha", "Perfil", "Departamento", "Ativo", "Data Cadastro"],
    "Tarefas": [
        "ID", "Tarefa", "Descrição", "Departamento", "Projeto", "Responsavel",
        "Periodicidade", "Obrigatoria", "Prioridade", "Dependencia",
        "Prazo Limite", "Data de Inicio", "Status", "Concluído Por",
        "Data Conclusão", "Observação", "Ativa"
    ],
    "Projetos": [
        "ID Projeto", "Projeto", "Objetivo", "Departamento", "Responsavel",
        "Data de Inicio", "Prazo Final", "Status", "%", "Proxima Etapa",
        "Observação", "Ativo"
    ],
    "Historico": ["ID Histórico", "ID Tarefa", "Tarefa", "Usuário", "Data", "Hora", "Status", "Observação"],
    "Calendario": ["Data", "ID Tarefa", "Departamento", "Status", "Prioridade"],
    "Configuracoes": ["Tipo", "Valor", "Ativo"],
}

# ============================================================
# CSS
# ============================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #111827 60%, #020617 100%);
}

section[data-testid="stSidebar"] * {
    color: #f8fafc !important;
}

div[data-testid="stSidebarContent"] {
    padding-top: 1.2rem;
}

.main-title {
    font-size: 34px;
    font-weight: 800;
    color: #111827;
    margin-bottom: 0;
}

.subtitle {
    color: #64748b;
    font-size: 15px;
    margin-top: 2px;
}

.metric-card {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 22px;
    padding: 22px;
    box-shadow: 0 8px 28px rgba(15, 23, 42, 0.06);
    min-height: 135px;
}

.metric-label {
    color: #334155;
    font-size: 15px;
    font-weight: 700;
}

.metric-value {
    font-size: 36px;
    font-weight: 800;
    margin-top: 8px;
}

.metric-detail {
    font-size: 13px;
    margin-top: 4px;
}

.panel {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 24px;
    padding: 22px;
    box-shadow: 0 8px 28px rgba(15, 23, 42, 0.05);
    margin-top: 18px;
}

.panel-title {
    font-size: 22px;
    font-weight: 800;
    color: #111827;
    margin-bottom: 18px;
}

.highlight-panel {
    background: linear-gradient(135deg, #eef2ff 0%, #f5f3ff 100%);
    border: 1px solid #c7d2fe;
    border-radius: 24px;
    padding: 22px;
    box-shadow: 0 8px 28px rgba(79, 70, 229, 0.08);
    margin-top: 18px;
}

.task-card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 18px;
    margin-bottom: 14px;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
}

.task-title {
    font-size: 18px;
    font-weight: 800;
    color: #111827;
}

.task-desc {
    color: #64748b;
    font-size: 14px;
    margin-top: 6px;
}

.tag {
    display: inline-block;
    padding: 6px 10px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 700;
    margin-right: 6px;
    margin-top: 8px;
}

.tag-blue { background:#dbeafe; color:#1d4ed8; }
.tag-green { background:#dcfce7; color:#15803d; }
.tag-red { background:#fee2e2; color:#b91c1c; }
.tag-yellow { background:#fef3c7; color:#b45309; }
.tag-purple { background:#ede9fe; color:#6d28d9; }
.tag-gray { background:#f1f5f9; color:#475569; }

.stButton > button {
    border-radius: 14px;
    min-height: 44px;
    font-weight: 700;
    border: 1px solid #d1d5db;
}

.stButton > button:hover {
    border-color: #7c3aed;
    color: #7c3aed;
}

input, textarea, select {
    border-radius: 12px !important;
}

.footer-box {
    background: #eef2ff;
    color: #3730a3;
    border: 1px solid #c7d2fe;
    border-radius: 18px;
    padding: 16px 20px;
    margin-top: 20px;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def txt(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def is_active(v):
    t = txt(v).lower()
    return t not in ["não", "nao", "n", "false", "0", "inativo", "inativa", "arquivada", "arquivado"]

def parse_date(v):
    if pd.isna(v) or txt(v) == "":
        return None
    try:
        return pd.to_datetime(v, dayfirst=True).date()
    except Exception:
        return None

def done(row):
    return txt(row.get("Status")).lower() in ["concluída", "concluida", "feito", "finalizada"]

def next_id(df, col):
    if df.empty or col not in df.columns:
        return 1
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    return int(s.max()) + 1 if not s.empty else 1

def ensure_cols(df, cols):
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]

def classify(row):
    if done(row):
        d = parse_date(row.get("Data Conclusão"))
        if d == date.today():
            return "Concluída hoje"
        return "Concluída"

    if not is_active(row.get("Ativa", "Sim")):
        return "Arquivada"

    data_ini = parse_date(row.get("Data de Inicio"))
    periodicidade = txt(row.get("Periodicidade")).lower()

    if data_ini and data_ini > date.today():
        return "Futura"

    if data_ini and data_ini < date.today() and periodicidade in ["unica", "única", "pontual", ""]:
        return "Pendente anterior"

    return "Hoje"

def periodic_today(row):
    if done(row) or not is_active(row.get("Ativa", "Sim")):
        return False

    data_ini = parse_date(row.get("Data de Inicio"))
    per = txt(row.get("Periodicidade")).lower()

    if data_ini and data_ini > date.today():
        return False

    if per in ["diario", "diária", "diaria", "diário", "todo dia", ""]:
        return True

    if per in ["semanal", "semana"]:
        return True if not data_ini else (date.today() - data_ini).days % 7 == 0

    if per in ["mensal", "mês", "mes"]:
        return True if not data_ini else date.today().day == data_ini.day

    if per in ["unica", "única", "pontual"]:
        return data_ini == date.today()

    return True

def belongs_to_user(row, user):
    resp = txt(row.get("Responsavel")).lower()
    u = txt(user).lower()
    if not resp or not u:
        return False

    nomes = [x.strip().lower() for x in resp.replace("/", ",").replace(";", ",").split(",") if x.strip()]
    return u in nomes or u in resp

def get_departamentos(data):
    tarefas = data.get("Tarefas", pd.DataFrame())
    if tarefas.empty or "Departamento" not in tarefas.columns:
        return []

    departamentos = sorted([
        txt(x)
        for x in tarefas["Departamento"].dropna().unique()
        if txt(x)
    ])

    return departamentos

# ============================================================
# DADOS
# ============================================================

@st.cache_data(show_spinner=False)
def load_data(path):
    path = Path(path)
    data = {}

    if not path.exists():
        for aba, cols in ABAS.items():
            data[aba] = pd.DataFrame(columns=cols)
        return data

    xls = pd.ExcelFile(path)
    for aba, cols in ABAS.items():
        if aba in xls.sheet_names:
            df = pd.read_excel(path, sheet_name=aba)
        else:
            df = pd.DataFrame(columns=cols)
        data[aba] = ensure_cols(df, cols)

    return data

def save_data(path, data):
    path = Path(path)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        for aba, cols in ABAS.items():
            ensure_cols(data.get(aba, pd.DataFrame(columns=cols)), cols).to_excel(writer, sheet_name=aba, index=False)
    st.cache_data.clear()

def get_path():
    if "excel_path" not in st.session_state:
        st.session_state["excel_path"] = ARQUIVO_EXCEL
    return st.session_state["excel_path"]

# ============================================================
# COMPONENTES VISUAIS
# ============================================================

def header(title, subtitle):
    col1, col2 = st.columns([4, 1.5])
    with col1:
        st.markdown(f"<div class='main-title'>{title}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='subtitle'>{subtitle}</div>", unsafe_allow_html=True)
    with col2:
        hoje = datetime.now()
        st.markdown(
            f"""
            <div style='text-align:right; color:#0f172a; font-weight:800; margin-top:8px;'>
                📅 {hoje.strftime('%d/%m/%Y')}
            </div>
            <div style='text-align:right; color:#64748b; font-size:14px;'>
                {hoje.strftime('%H:%M')}
            </div>
            """,
            unsafe_allow_html=True
        )

def metric_card(label, value, detail, color):
    st.markdown(
        f"""
        <div class='metric-card'>
            <div class='metric-label'>{label}</div>
            <div class='metric-value' style='color:{color};'>{value}</div>
            <div class='metric-detail' style='color:{color};'>{detail}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

def sidebar(data):
    st.sidebar.markdown("## ✅ Agenda\n### Operacional")
    st.sidebar.divider()

    usuarios = data["Usuarios"].copy()
    nomes = usuarios[usuarios["Ativo"].apply(is_active)]["Nome"].dropna().astype(str).str.strip().tolist()
    nomes = sorted([n for n in nomes if n]) or ["Paula"]

    user = st.sidebar.selectbox("Usuário logado", nomes)
    st.sidebar.caption("Todos visualizam tudo. As tarefas do usuário logado aparecem em destaque.")

    st.sidebar.divider()

    departamentos = get_departamentos(data)

    menu = ["Dashboard", "Minhas tarefas"] + departamentos + [
        "Projetos",
        "Calendário",
        "Cadastro de tarefas",
        "Histórico"
    ]

    page = st.sidebar.radio("Navegação", menu)

    st.sidebar.divider()
    if st.sidebar.button("🔄 Atualizar dados"):
        st.cache_data.clear()
        st.rerun()

    st.sidebar.caption("Versão 1.2.0")
    return user, page, departamentos

def prepared_tasks(data):
    tarefas = data["Tarefas"].copy()
    tarefas = tarefas[tarefas["Ativa"].apply(is_active)]
    tarefas["Classificacao"] = tarefas.apply(classify, axis=1)
    return tarefas

def task_card(row, data, user, prefix):
    tarefas = data["Tarefas"].copy()
    hist = data["Historico"].copy()

    idt = row.get("ID")
    titulo = txt(row.get("Tarefa"))
    desc = txt(row.get("Descrição"))
    depto = txt(row.get("Departamento"))
    resp = txt(row.get("Responsavel"))
    prioridade = txt(row.get("Prioridade")) or "Normal"
    projeto = txt(row.get("Projeto"))
    periodicidade = txt(row.get("Periodicidade"))
    status = classify(row)
    minha = belongs_to_user(row, user)

    tag_status = "tag-green" if "Concluída" in status else "tag-red" if "anterior" in status else "tag-yellow"

    with st.container():
        st.markdown("<div class='task-card'>", unsafe_allow_html=True)
        c1, c2 = st.columns([4.8, 1.25])

        with c1:
            st.markdown(f"<div class='task-title'>{'⭐ ' if minha else ''}{titulo}</div>", unsafe_allow_html=True)
            if desc:
                st.markdown(f"<div class='task-desc'>{desc}</div>", unsafe_allow_html=True)

            tags = f"<span class='tag {tag_status}'>{status}</span>"
            if minha:
                tags += "<span class='tag tag-purple'>Minha tarefa</span>"
            if depto:
                tags += f"<span class='tag tag-blue'>{depto}</span>"
            if projeto:
                tags += f"<span class='tag tag-purple'>Projeto: {projeto}</span>"
            if prioridade:
                tags += f"<span class='tag tag-yellow'>{prioridade}</span>"
            st.markdown(tags, unsafe_allow_html=True)

            st.caption(f"Responsável: {resp or '-'} | Periodicidade: {periodicidade or '-'}")

        with c2:
            if done(row):
                st.success("✅ Concluída")
                if st.button("Reabrir", key=f"reopen_{prefix}_{idt}"):
                    idxs = tarefas.index[tarefas["ID"].astype(str) == str(idt)].tolist()
                    if idxs:
                        i = idxs[0]
                        tarefas.at[i, "Status"] = "Pendente"
                        tarefas.at[i, "Concluído Por"] = ""
                        tarefas.at[i, "Data Conclusão"] = ""
                        data["Tarefas"] = tarefas
                        save_data(get_path(), data)
                        st.rerun()
            else:
                if st.button("✅ Concluir", key=f"done_{prefix}_{idt}"):
                    idxs = tarefas.index[tarefas["ID"].astype(str) == str(idt)].tolist()
                    if not idxs:
                        st.error("Tarefa não localizada no Excel.")
                    else:
                        now = datetime.now()
                        i = idxs[0]
                        tarefas.at[i, "Status"] = "Concluída"
                        tarefas.at[i, "Concluído Por"] = user
                        tarefas.at[i, "Data Conclusão"] = now.strftime("%d/%m/%Y %H:%M:%S")

                        novo = {
                            "ID Histórico": next_id(hist, "ID Histórico"),
                            "ID Tarefa": idt,
                            "Tarefa": titulo,
                            "Usuário": user,
                            "Data": now.strftime("%d/%m/%Y"),
                            "Hora": now.strftime("%H:%M:%S"),
                            "Status": "Concluída",
                            "Observação": "",
                        }

                        hist = pd.concat([hist, pd.DataFrame([novo])], ignore_index=True)

                        data["Tarefas"] = tarefas
                        data["Historico"] = hist
                        save_data(get_path(), data)
                        st.success("Tarefa concluída.")
                        st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# PÁGINAS
# ============================================================

def dashboard(data, user):
    header("Dashboard", "Visão geral das atividades")

    tarefas = prepared_tasks(data)

    minhas = tarefas[tarefas.apply(lambda r: belongs_to_user(r, user), axis=1)]
    minhas_abertas = minhas[~minhas.apply(done, axis=1)]

    pend = tarefas[tarefas["Classificacao"] == "Pendente anterior"]
    hoje = tarefas[(tarefas.apply(periodic_today, axis=1)) & (~tarefas.apply(done, axis=1))]
    concl_hoje = tarefas[tarefas["Classificacao"] == "Concluída hoje"]

    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("⭐ Minhas pendentes", len(minhas_abertas), f"{user}", "#7c3aed")
    with c2: metric_card("Tarefas de hoje", len(hoje), f"{len(hoje)} pendente(s)", "#f59e0b")
    with c3: metric_card("Pendências anteriores", len(pend), "Atenção" if len(pend) else "Sem pendências", "#dc2626")
    with c4: metric_card("Concluídas hoje", len(concl_hoje), f"{len(concl_hoje)} concluída(s)", "#16a34a")

    st.markdown("<div class='highlight-panel'><div class='panel-title'>⭐ Minhas tarefas em destaque</div>", unsafe_allow_html=True)
    if minhas_abertas.empty:
        st.success("Você não possui tarefas pendentes no momento.")
    else:
        for _, row in minhas_abertas.iterrows():
            task_card(row, data, user, "minhas_dash")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns([2,1,1,1])
    with f1:
        busca = st.text_input("🔎 Buscar tarefa", placeholder="Digite uma palavra...")
    with f2:
        deptos = ["Todos"] + sorted([x for x in tarefas["Departamento"].dropna().astype(str).unique() if x])
        depto = st.selectbox("Departamento", deptos)
    with f3:
        resps = ["Todos"] + sorted([x for x in tarefas["Responsavel"].dropna().astype(str).unique() if x])
        resp = st.selectbox("Responsável", resps)
    with f4:
        status = st.selectbox("Status", ["Todos", "Pendente anterior", "Hoje", "Concluída hoje", "Concluída", "Futura"])

    filtradas = tarefas.copy()
    if busca:
        b = busca.lower()
        filtradas = filtradas[
            filtradas["Tarefa"].astype(str).str.lower().str.contains(b, na=False) |
            filtradas["Descrição"].astype(str).str.lower().str.contains(b, na=False)
        ]
    if depto != "Todos":
        filtradas = filtradas[filtradas["Departamento"].astype(str) == depto]
    if resp != "Todos":
        filtradas = filtradas[filtradas["Responsavel"].astype(str) == resp]
    if status != "Todos":
        filtradas = filtradas[filtradas["Classificacao"] == status]

    colA, colB = st.columns(2)

    with colA:
        st.markdown("<div class='panel'><div class='panel-title'>🔴 Pendências anteriores</div>", unsafe_allow_html=True)
        lista = filtradas[filtradas["Classificacao"] == "Pendente anterior"]
        if lista.empty:
            st.success("Nenhuma pendência anterior.")
        else:
            for _, row in lista.iterrows():
                task_card(row, data, user, "pend")
        st.markdown("</div>", unsafe_allow_html=True)

    with colB:
        st.markdown("<div class='panel'><div class='panel-title'>🟡 Agenda de hoje</div>", unsafe_allow_html=True)
        lista = filtradas[(filtradas.apply(periodic_today, axis=1)) & (~filtradas.apply(done, axis=1))]
        if lista.empty:
            st.info("Nenhuma tarefa pendente para hoje.")
        else:
            for _, row in lista.iterrows():
                task_card(row, data, user, "hoje")
        st.markdown("</div>", unsafe_allow_html=True)

def minhas_tarefas_page(data, user):
    header("Minhas tarefas", f"Atividades atribuídas a {user}")

    tarefas = prepared_tasks(data)
    minhas = tarefas[tarefas.apply(lambda r: belongs_to_user(r, user), axis=1)]

    c1, c2, c3 = st.columns(3)
    with c1: metric_card("Minhas abertas", len(minhas[~minhas.apply(done, axis=1)]), "Pendentes", "#7c3aed")
    with c2: metric_card("Minhas concluídas", len(minhas[minhas.apply(done, axis=1)]), "Histórico visual", "#16a34a")
    with c3: metric_card("Total atribuídas", len(minhas), "Todas", "#2563eb")

    st.markdown("<div class='highlight-panel'><div class='panel-title'>⭐ Prioridade do usuário logado</div>", unsafe_allow_html=True)

    abertas = minhas[~minhas.apply(done, axis=1)]
    if abertas.empty:
        st.success("Você não possui tarefas pendentes.")
    else:
        for _, row in abertas.iterrows():
            task_card(row, data, user, "minhas_page")

    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander("Ver minhas tarefas concluídas"):
        concluidas = minhas[minhas.apply(done, axis=1)]
        if concluidas.empty:
            st.info("Nenhuma tarefa concluída.")
        else:
            st.dataframe(concluidas, use_container_width=True, hide_index=True)

def departamento_page(data, user, depto):
    header(depto, f"Tarefas do departamento {depto}")
    tarefas = prepared_tasks(data)
    tarefas = tarefas[tarefas["Departamento"].astype(str).str.lower().str.strip() == depto.lower()]

    busca = st.text_input("🔎 Buscar tarefa")
    somente_minhas = st.toggle("Mostrar somente minhas tarefas neste departamento", value=False)

    if somente_minhas:
        tarefas = tarefas[tarefas.apply(lambda r: belongs_to_user(r, user), axis=1)]

    if busca:
        b = busca.lower()
        tarefas = tarefas[
            tarefas["Tarefa"].astype(str).str.lower().str.contains(b, na=False) |
            tarefas["Descrição"].astype(str).str.lower().str.contains(b, na=False)
        ]

    if tarefas.empty:
        st.info("Nenhuma tarefa encontrada.")
    else:
        for _, row in tarefas.iterrows():
            task_card(row, data, user, f"depto_{depto}")

def projetos_page(data, user):
    header("Projetos", "Acompanhamento dos projetos internos")
    projetos = data["Projetos"].copy()
    projetos = projetos[projetos["Ativo"].apply(is_active)] if not projetos.empty else projetos

    if projetos.empty:
        st.info("Nenhum projeto cadastrado.")
        return

    for _, row in projetos.iterrows():
        pct = pd.to_numeric(row.get("%"), errors="coerce")
        pct = 0 if pd.isna(pct) else float(pct)
        with st.container(border=True):
            c1, c2 = st.columns([4,1])
            with c1:
                st.markdown(f"### {txt(row.get('Projeto'))}")
                st.write(txt(row.get("Objetivo")))
                st.caption(f"Responsável: {txt(row.get('Responsavel')) or '-'} | Prazo: {txt(row.get('Prazo Final')) or '-'} | Status: {txt(row.get('Status')) or '-'}")
                st.progress(min(max(pct/100, 0), 1))
            with c2:
                st.metric("Concluído", f"{pct:.0f}%")

def cadastro_page(data, user):
    header("Cadastro de tarefas", "Inclua novas atividades no sistema")

    tarefas = data["Tarefas"].copy()
    usuarios = data["Usuarios"].copy()
    projetos = data["Projetos"].copy()

    deptos = get_departamentos(data) or ["Financeiro", "Controladoria", "Contabilidade", "Projetos"]
    responsaveis = [""] + sorted([x for x in usuarios["Nome"].dropna().astype(str).unique() if x])
    projetos_l = [""] + sorted([x for x in projetos["Projeto"].dropna().astype(str).unique() if x])
    tarefas_l = [""] + sorted([x for x in tarefas["Tarefa"].dropna().astype(str).unique() if x])

    with st.form("nova_tarefa"):
        c1, c2 = st.columns(2)
        with c1:
            tarefa = st.text_input("Tarefa")
            desc = st.text_area("Descrição")
            depto = st.selectbox("Departamento", deptos)
            projeto = st.selectbox("Projeto", projetos_l)
            resp = st.selectbox("Responsável", responsaveis)
        with c2:
            per = st.selectbox("Periodicidade", ["Diario", "Semanal", "Mensal", "Unica"])
            obrig = st.selectbox("Obrigatória", ["Não", "Sim"])
            prio = st.selectbox("Prioridade", ["Normal", "Alta", "Crítica", "Baixa"])
            dep = st.selectbox("Dependência", tarefas_l)
            data_ini = st.date_input("Data de início", value=date.today())
            prazo = st.time_input("Prazo limite", value=time(18,0))
        obs = st.text_area("Observação")
        ok = st.form_submit_button("Salvar tarefa")

    if ok:
        if not tarefa:
            st.error("Informe a tarefa.")
        else:
            novo = {
                "ID": next_id(tarefas, "ID"),
                "Tarefa": tarefa,
                "Descrição": desc,
                "Departamento": depto,
                "Projeto": projeto,
                "Responsavel": resp,
                "Periodicidade": per,
                "Obrigatoria": obrig,
                "Prioridade": prio,
                "Dependencia": dep,
                "Prazo Limite": prazo.strftime("%H:%M:%S"),
                "Data de Inicio": data_ini.strftime("%d/%m/%Y"),
                "Status": "Pendente",
                "Concluído Por": "",
                "Data Conclusão": "",
                "Observação": obs,
                "Ativa": "Sim",
            }
            data["Tarefas"] = pd.concat([tarefas, pd.DataFrame([novo])], ignore_index=True)
            save_data(get_path(), data)
            st.success("Tarefa cadastrada.")
            st.rerun()

def calendario_page(data, user):
    header("Calendário", "Visão operacional por data")
    tarefas = prepared_tasks(data)
    ref = st.date_input("Data", value=date.today())

    st.dataframe(
        tarefas[["ID", "Tarefa", "Departamento", "Responsavel", "Periodicidade", "Prioridade", "Status", "Data de Inicio"]],
        use_container_width=True,
        hide_index=True
    )

def historico_page(data):
    header("Histórico", "Registro das conclusões")
    hist = data["Historico"].copy()
    if hist.empty:
        st.info("Ainda não há histórico.")
    else:
        st.dataframe(hist.sort_values("ID Histórico", ascending=False), use_container_width=True, hide_index=True)

# ============================================================
# MAIN
# ============================================================

def main():
    data = load_data(get_path())
    user, page, departamentos = sidebar(data)

    if page == "Dashboard":
        dashboard(data, user)
    elif page == "Minhas tarefas":
        minhas_tarefas_page(data, user)
    elif page in departamentos:
        departamento_page(data, user, page)
    elif page == "Projetos":
        projetos_page(data, user)
    elif page == "Cadastro de tarefas":
        cadastro_page(data, user)
    elif page == "Calendário":
        calendario_page(data, user)
    elif page == "Histórico":
        historico_page(data)

if __name__ == "__main__":
    main()
