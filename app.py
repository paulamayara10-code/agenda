
# -*- coding: utf-8 -*-
from datetime import datetime, date, timedelta
from pathlib import Path
import shutil
import unicodedata

import pandas as pd
import streamlit as st

from database import (
    DB_PATH,
    init_db,
    banco_tem_dados,
    listar_usuarios,
    listar_departamentos,
    listar_tarefas,
    listar_historico,
    listar_projetos,
    criar_usuario,
    criar_tarefa,
    atualizar_status,
    adicionar_observacao,
    editar_tarefa,
    arquivar_tarefa,
    reativar_tarefa,
    exportar_excel,
)
from migrar_excel_para_sqlite import migrar_excel_para_sqlite


st.set_page_config(
    page_title="Agenda Operacional SQLite",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded"
)

init_db(DB_PATH)

# Migração automática na primeira execução
if not banco_tem_dados(DB_PATH) and Path("Agenda.xlsx").exists():
    ok_migracao, msg_migracao = migrar_excel_para_sqlite("Agenda.xlsx", DB_PATH, somente_se_vazio=True)
else:
    ok_migracao, msg_migracao = False, ""


st.markdown("""
<style>
html, body, [class*="css"] { font-family: 'Inter', Arial, sans-serif; }
.stApp { background: #f5f7fb; }
div.block-container { padding-top: 1.5rem; padding-bottom: 4rem; max-width: 1450px; }
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #0f172a 0%, #111827 70%, #1e1b4b 100%); }
section[data-testid="stSidebar"] * { color: #ffffff !important; }
section[data-testid="stSidebar"] div[data-baseweb="select"] * { color: #0f172a !important; }
.app-title { font-size: 38px; font-weight: 900; color: #0f172a; letter-spacing: -0.04em; margin-bottom: 0; }
.app-subtitle { color: #64748b; font-size: 15px; margin-top: 4px; margin-bottom: 22px; }
.hero { background: linear-gradient(135deg, #0f172a 0%, #312e81 45%, #7c3aed 100%); color: white; border-radius: 30px; padding: 28px 32px; box-shadow: 0 22px 65px rgba(49,46,129,0.25); margin-bottom: 24px; }
.hero h1 { font-size: 32px; margin: 0; font-weight: 900; }
.hero p { color: #ddd6fe; margin: 8px 0 0 0; }
.metric-card { background: white; border: 1px solid #e2e8f0; border-radius: 24px; padding: 22px; box-shadow: 0 14px 35px rgba(15,23,42,0.06); min-height: 130px; }
.metric-label { font-size: 13px; color: #64748b; text-transform: uppercase; font-weight: 800; letter-spacing: .03em; }
.metric-value { font-size: 38px; font-weight: 900; margin-top: 8px; }
.metric-detail { font-size: 13px; margin-top: 4px; }
.panel { background: white; border: 1px solid #e2e8f0; border-radius: 28px; padding: 22px; box-shadow: 0 16px 42px rgba(15,23,42,0.06); margin-top: 18px; }
.panel-title { font-size: 22px; font-weight: 900; color: #0f172a; margin-bottom: 16px; }
.task-title { font-size: 18px; font-weight: 900; color: #0f172a; }
.task-title-done { font-size: 18px; font-weight: 800; color: #94a3b8; text-decoration: line-through; }
.tag { display: inline-block; padding: 6px 10px; border-radius: 999px; font-size: 12px; font-weight: 800; margin-right: 6px; margin-top: 8px; }
.tag-blue { background:#dbeafe; color:#1d4ed8; }
.tag-green { background:#dcfce7; color:#15803d; }
.tag-red { background:#fee2e2; color:#b91c1c; }
.tag-yellow { background:#fef3c7; color:#b45309; }
.tag-purple { background:#ede9fe; color:#6d28d9; }
.tag-orange { background:#ffedd5; color:#c2410c; }
.tag-gray { background:#f1f5f9; color:#475569; }
.stButton > button { border-radius: 14px; font-weight: 800; min-height: 42px; }

.data-global-note {
    font-size: 13px;
    color: #64748b;
}
</style>
""", unsafe_allow_html=True)


def txt(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def normalizar_departamento(v):
    t = txt(v)
    mapa = {
        "contas a receber": "Contas a Receber",
        "contas a pagar": "Contas a Pagar",
        "contabilidade": "Contabilidade",
        "controladoria": "Controladoria",
        "tesouraria": "Tesouraria",
        "financeiro": "Financeiro",
        "projetos": "Projetos",
    }
    return mapa.get(" ".join(t.lower().split()), t.strip().title())


def normalizar_periodicidade(v):
    t = txt(v).lower().strip()
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode("ascii")
    t = " ".join(t.split())

    if t in ["diario", "diaria", "todo dia", "todos os dias", "dia", "daily", ""]:
        return "diario"
    if t in ["semanal", "semana", "weekly", "toda semana"]:
        return "semanal"
    if t in ["mensal", "mes", "mensalmente", "monthly"]:
        return "mensal"
    if t in ["unica", "unico", "pontual", "uma vez"]:
        return "unica"
    return t


def parse_data(v):
    if not txt(v):
        return None
    try:
        return pd.to_datetime(v, dayfirst=True).date()
    except Exception:
        return None


def hoje():
    return date.today()


def eh_dia_util(data_ref):
    return data_ref.weekday() < 5


def data_referencia_global():
    if "data_referencia" not in st.session_state:
        st.session_state["data_referencia"] = hoje()
    return st.session_state["data_referencia"]


def dia_util_anterior(data_ref):
    data_ant = data_ref - timedelta(days=1)
    while not eh_dia_util(data_ant):
        data_ant = data_ant - timedelta(days=1)
    return data_ant


def tarefa_concluida_em_data(id_tarefa, data_ref):
    hist = listar_historico()
    if hist.empty:
        return False

    data_str = data_ref.strftime("%d/%m/%Y")

    filtro = (
        (hist["id_tarefa"].astype(str) == str(id_tarefa))
        & (hist["status"].astype(str).str.lower().isin(["concluída", "concluida"]))
        & (hist["data"].astype(str) == data_str)
    )

    return filtro.any()


def eh_pendencia_anterior(row, ref=None):
    """
    Regra gerencial:
    - tarefa única com data anterior e não concluída = pendência anterior;
    - tarefa diária vira pendência anterior se não houve conclusão no último dia útil anterior;
    - fim de semana não gera pendência.
    """
    if ref is None:
        ref = data_referencia_global()

    if not eh_dia_util(ref):
        return False

    if is_concluida(row, ref):
        return False

    per = normalizar_periodicidade(row.get("periodicidade"))
    data_ini = parse_data(row.get("data_inicio"))
    id_tarefa = row.get("id")

    if per == "diario":
        ant = dia_util_anterior(ref)

        if data_ini and data_ini > ant:
            return False

        return not tarefa_concluida_em_data(id_tarefa, ant)

    if per == "unica":
        return bool(data_ini and data_ini < ref and not is_concluida(row, ref))

    return False


def is_concluida(row, ref=None):
    if ref is None:
        ref = hoje()
    status = txt(row.get("status")).lower()
    if status not in ["concluída", "concluida", "feito", "finalizada"]:
        return False
    per = txt(row.get("periodicidade")).lower()
    data_conc = parse_data(row.get("data_conclusao"))
    if per in ["unica", "única", "pontual"]:
        return True
    return data_conc == ref


def periodic_on_date(row, ref=None):
    """
    Regra com dias úteis:
    - Tarefa diária aparece apenas em dias úteis.
    - Tarefa sem data de início é considerada ativa em dias úteis.
    - Semanal/mensal respeitam a data de início.
    - Única em fim de semana aparece no próximo dia útil.
    """
    if ref is None:
        ref = data_referencia_global()

    if is_concluida(row, ref):
        return False

    if not eh_dia_util(ref):
        return False

    data_ini = parse_data(row.get("data_inicio"))
    per = normalizar_periodicidade(row.get("periodicidade"))

    if per == "diario":
        if data_ini and data_ini > ref:
            return False
        return True

    if not data_ini:
        return True

    if data_ini > ref:
        return False

    if per == "semanal":
        return (ref - data_ini).days % 7 == 0

    if per == "mensal":
        return ref.day == data_ini.day

    if per == "unica":
        data_exec = data_ini
        while not eh_dia_util(data_exec):
            data_exec = data_exec + timedelta(days=1)
        return data_exec == ref

    return True

def classificar(row, ref=None):
    if ref is None:
        ref = data_referencia_global()

    if is_concluida(row, ref):
        return "Concluída na data"

    status = txt(row.get("status")).lower()
    if status in ["em andamento", "andamento", "iniciada", "iniciado"]:
        return "Em andamento"

    if not eh_dia_util(ref):
        return "Fim de semana"

    data_ini = parse_data(row.get("data_inicio"))
    per = normalizar_periodicidade(row.get("periodicidade"))

    if per == "diario":
        if data_ini and data_ini > ref:
            return "Futura"
        return "Hoje" if ref == hoje() else "Prevista"

    if not data_ini:
        return "Hoje" if ref == hoje() else "Prevista"

    if data_ini > ref:
        return "Futura"

    if data_ini < ref and per == "unica":
        return "Pendente anterior"

    if periodic_on_date(row, ref):
        return "Hoje" if ref == hoje() else "Prevista"

    return "Futura"

def pertence_usuario(row, user):
    resp = txt(row.get("responsavel")).lower()
    u = txt(user).lower()
    if not resp or not u:
        return False
    nomes = [x.strip().lower() for x in resp.replace("/", ",").replace(";", ",").split(",")]
    return u in nomes or u in resp


def ultima_observacao(id_tarefa):
    hist = listar_historico()
    if hist.empty:
        return ""
    filtro = (
        (hist["id_tarefa"].astype(str) == str(id_tarefa))
        & (hist["status"].astype(str).str.lower() == "observação")
    )
    obs = hist[filtro]
    if obs.empty:
        return ""
    linha = obs.iloc[0]
    return f"{linha.get('data','')} {linha.get('hora','')} - {linha.get('usuario','')}: {linha.get('observacao','')}"


def preparar_tarefas(ref=None):
    if ref is None:
        ref = data_referencia_global()
    df = listar_tarefas(ativas=True)
    if df.empty:
        return df
    df["departamento"] = df["departamento"].apply(normalizar_departamento)
    df["classificacao"] = df.apply(lambda r: classificar(r, ref), axis=1)
    return df


def metric_card(label, value, detail, color):
    st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>{label}</div>
            <div class='metric-value' style='color:{color};'>{value}</div>
            <div class='metric-detail' style='color:{color};'>{detail}</div>
        </div>
    """, unsafe_allow_html=True)


def header(title, subtitle):
    col1, col2 = st.columns([4, 1.5])
    with col1:
        st.markdown(f"<div class='app-title'>{title}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='app-subtitle'>{subtitle}</div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div style='text-align:right;font-weight:800;color:#0f172a;'>📅 {datetime.now().strftime('%d/%m/%Y • %H:%M')}</div>", unsafe_allow_html=True)


def seletor_data_referencia():
    """
    Barra global de data.
    Todas as telas usam a mesma data salva no session_state.
    """
    if "data_referencia" not in st.session_state:
        st.session_state["data_referencia"] = hoje()

    st.markdown("<div class='panel' style='padding:16px 20px; margin-top:0;'>", unsafe_allow_html=True)

    c0, c1, c2, c3, c4 = st.columns([1.1, 1.2, 0.8, 0.8, 3])

    with c0:
        st.markdown("#### 📅 Data global")

    with c1:
        data_sel = st.date_input(
            "Data de referência",
            value=st.session_state["data_referencia"],
            label_visibility="collapsed",
            key="data_global_input"
        )
        st.session_state["data_referencia"] = data_sel

    with c2:
        if st.button("◀ Dia anterior", use_container_width=True):
            st.session_state["data_referencia"] = st.session_state["data_referencia"] - timedelta(days=1)
            st.rerun()

    with c3:
        if st.button("Hoje", use_container_width=True):
            st.session_state["data_referencia"] = hoje()
            st.rerun()

    with c4:
        c4a, c4b = st.columns([0.9, 2])
        with c4a:
            if st.button("Próximo dia ▶", use_container_width=True):
                st.session_state["data_referencia"] = st.session_state["data_referencia"] + timedelta(days=1)
                st.rerun()
        with c4b:
            data_ref = st.session_state["data_referencia"]
            if eh_dia_util(data_ref):
                st.success(f"Dia útil selecionado: {data_ref.strftime('%d/%m/%Y')}")
            else:
                st.warning("Sábado/domingo selecionado. Não gera tarefas recorrentes.")

    st.markdown("</div>", unsafe_allow_html=True)

    return st.session_state["data_referencia"]



def sidebar():
    st.sidebar.markdown("## ✅ Agenda")
    st.sidebar.markdown("### Operacional SQLite")
    st.sidebar.divider()
    usuarios = listar_usuarios()
    nomes = usuarios["nome"].dropna().tolist() if not usuarios.empty else []
    if not nomes:
        criar_usuario("Paula", "Administradora")
        nomes = ["Paula"]
    user = st.sidebar.selectbox("Usuário logado", nomes)
    deps = listar_departamentos()
    deps = [normalizar_departamento(d) for d in deps]
    deps = sorted(list(dict.fromkeys(deps)))
    menu = ["Dashboard", "Coordenação", "Minhas tarefas"] + deps + [
        "Projetos", "Calendário", "Cadastro de tarefas", "Editar tarefas",
        "Pendências com observação", "Tarefas arquivadas", "Histórico", "Exportar Excel", "Admin SQLite"
    ]
    page = st.sidebar.radio("Navegação", menu)
    return user, page, deps


def app_frame(title, subtitle):
    header(title, subtitle)
    return seletor_data_referencia()


def task_card(row, user, prefix="task"):
    idt = int(row.get("id"))
    titulo = txt(row.get("tarefa"))
    desc = txt(row.get("descricao"))
    depto = txt(row.get("departamento"))
    resp = txt(row.get("responsavel"))
    prioridade = txt(row.get("prioridade")) or "Normal"
    projeto = txt(row.get("projeto"))
    periodicidade = txt(row.get("periodicidade"))
    status = txt(row.get("status"))
    minha = pertence_usuario(row, user)
    done = is_concluida(row, data_referencia_global())
    with st.container(border=True):
        c1, c2 = st.columns([4.6, 1.35])
        with c1:
            cls = "task-title-done" if done else "task-title"
            st.markdown(f"<div class='{cls}'>{'⭐ ' if minha else ''}{titulo}</div>", unsafe_allow_html=True)
            if desc:
                st.caption(desc)
            tags = ""
            if status:
                tags += f"<span class='tag tag-purple'>{status}</span>"
            if minha:
                tags += "<span class='tag tag-blue'>Minha tarefa</span>"
            if depto:
                tags += f"<span class='tag tag-green'>{depto}</span>"
            if projeto:
                tags += f"<span class='tag tag-purple'>Projeto: {projeto}</span>"
            if prioridade:
                tags += f"<span class='tag tag-yellow'>{prioridade}</span>"
            st.markdown(tags, unsafe_allow_html=True)
            st.caption(f"Responsável: {resp or '-'} | Periodicidade: {periodicidade or '-'}")
            obs = ultima_observacao(idt)
            if obs:
                st.info(f"Última observação: {obs}")
        with c2:
            if done:
                st.success("✅ Concluída")
                if st.button("Reabrir", key=f"reopen_{prefix}_{idt}"):
                    atualizar_status(idt, "Pendente", user, "Tarefa reaberta")
                    st.rerun()
            else:
                if status.lower() != "em andamento":
                    if st.button("▶️ Iniciar", key=f"start_{prefix}_{idt}"):
                        atualizar_status(idt, "Em andamento", user, "Tarefa iniciada")
                        st.rerun()
                if st.button("✅ Concluir", key=f"done_{prefix}_{idt}"):
                    atualizar_status(idt, "Concluída", user, "Tarefa concluída")
                    st.rerun()
        with st.expander("💬 Observações / justificativa", expanded=False):
            observacao = st.text_area("Registrar observação", placeholder="Ex.: iniciado, aguardando documento, faltou retorno do banco...", key=f"obs_{prefix}_{idt}")
            if st.button("Salvar observação", key=f"save_obs_{prefix}_{idt}"):
                if adicionar_observacao(idt, user, observacao):
                    st.success("Observação registrada.")
                    st.rerun()
                else:
                    st.warning("Digite uma observação antes de salvar.")


def dashboard(user):
    ref = app_frame("Dashboard", "Visão geral das atividades")
    if ok_migracao and msg_migracao:
        st.success(msg_migracao)
    st.markdown(f"""
        <div class='hero'>
            <h1>Agenda operacional da equipe</h1>
            <p>Olá, {user}. Acompanhe prioridades, pendências e tarefas do dia em tempo real.</p>
        </div>
    """, unsafe_allow_html=True)
    tarefas = preparar_tarefas(ref)
    if tarefas.empty:
        st.info("Nenhuma tarefa cadastrada ainda. Verifique se o Agenda.xlsx está na mesma pasta do app.")
        return
    minhas = tarefas[tarefas.apply(lambda r: pertence_usuario(r, user), axis=1)]
    minhas_abertas = minhas[~minhas.apply(lambda r: is_concluida(r, ref), axis=1)]
    hoje_df = tarefas[(tarefas.apply(lambda r: periodic_on_date(r, ref), axis=1)) & (~tarefas.apply(lambda r: is_concluida(r, ref), axis=1))]
    pend = tarefas[tarefas.apply(lambda r: eh_pendencia_anterior(r, ref), axis=1)]
    concl = tarefas[tarefas["classificacao"] == "Concluída hoje"]
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("⭐ Minhas pendentes", len(minhas_abertas), user, "#7c3aed")
    with c2: metric_card("Tarefas de hoje", len(hoje_df), "Pendentes", "#f59e0b")
    with c3: metric_card("Pendências anteriores", len(pend), "Atenção", "#dc2626")
    with c4: metric_card("Concluídas hoje", len(concl), "Produtividade", "#16a34a")
    st.markdown("<div class='panel'><div class='panel-title'>⭐ Minhas tarefas em destaque</div>", unsafe_allow_html=True)
    if minhas_abertas.empty:
        st.success("Você não possui tarefas pendentes.")
    else:
        for _, row in minhas_abertas.iterrows():
            task_card(row, user, "dash_minhas")
    st.markdown("</div>", unsafe_allow_html=True)
    cA, cB = st.columns(2)
    with cA:
        st.markdown("<div class='panel'><div class='panel-title'>🔴 Pendências anteriores</div>", unsafe_allow_html=True)
        if pend.empty:
            st.success("Sem pendências anteriores.")
        else:
            for _, row in pend.iterrows():
                task_card(row, user, "dash_pend")
        st.markdown("</div>", unsafe_allow_html=True)
    with cB:
        st.markdown("<div class='panel'><div class='panel-title'>🟡 Agenda de hoje</div>", unsafe_allow_html=True)
        if hoje_df.empty:
            st.info("Sem tarefas para hoje.")
        else:
            for _, row in hoje_df.iterrows():
                task_card(row, user, "dash_hoje")
        st.markdown("</div>", unsafe_allow_html=True)


def coordenacao(user):
    ref = app_frame("Coordenação", "Central gerencial da operação")
    tarefas = preparar_tarefas(ref)
    hist = listar_historico()
    if tarefas.empty:
        st.info("Nenhuma tarefa cadastrada.")
        return
    abertas = tarefas[~tarefas.apply(lambda r: is_concluida(r, ref), axis=1)]
    atrasadas = abertas[abertas.apply(lambda r: eh_pendencia_anterior(r, ref), axis=1)]
    em_andamento = abertas[abertas["status"].astype(str).str.lower() == "em andamento"]
    sem_resp = abertas[abertas["responsavel"].astype(str).str.strip() == ""]
    concl = tarefas[tarefas["classificacao"] == "Concluída hoje"]
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: metric_card("🔴 Atrasadas", len(atrasadas), "Gargalos", "#dc2626")
    with c2: metric_card("▶️ Em andamento", len(em_andamento), "Iniciadas", "#f97316")
    with c3: metric_card("🟡 Abertas", len(abertas), "Pendentes", "#f59e0b")
    with c4: metric_card("✅ Concluídas", len(concl), "Hoje", "#16a34a")
    with c5: metric_card("👤 Sem responsável", len(sem_resp), "Ajustar", "#2563eb")
    st.markdown("<div class='panel'><div class='panel-title'>📊 Ranking operacional</div>", unsafe_allow_html=True)
    ranking = []
    for resp in sorted([x for x in tarefas["responsavel"].dropna().unique() if txt(x)]):
        base = tarefas[tarefas["responsavel"].astype(str).str.contains(str(resp), case=False, na=False)]
        abertas_r = base[~base.apply(lambda r: is_concluida(r, ref), axis=1)]
        atrasadas_r = abertas_r[abertas_r.apply(lambda r: eh_pendencia_anterior(r, ref), axis=1)]
        concl_r = hist[(hist["usuario"].astype(str).str.lower() == str(resp).lower()) & (hist["status"].astype(str).str.lower().isin(["concluída", "concluida"])) & (hist["data"].astype(str) == hoje().strftime("%d/%m/%Y"))] if not hist.empty else pd.DataFrame()
        score = max(0, min(100, 100 - len(atrasadas_r)*12 - len(abertas_r)*3 + len(concl_r)*2))
        ranking.append({"Usuário": resp, "Pendentes": len(abertas_r), "Atrasadas": len(atrasadas_r), "Concluídas Hoje": len(concl_r), "Score": score})
    if ranking:
        st.dataframe(pd.DataFrame(ranking), use_container_width=True, hide_index=True)
    else:
        st.info("Ainda não há responsáveis suficientes.")
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<div class='panel'><div class='panel-title'>🕘 Feed operacional</div>", unsafe_allow_html=True)
    if hist.empty:
        st.info("Ainda não há histórico.")
    else:
        for _, row in hist.head(15).iterrows():
            st.markdown(f"**{row.get('hora','')}** — {row.get('usuario','')} registrou **{row.get('status','')}** em “{row.get('tarefa','')}”")
            if txt(row.get("observacao")):
                st.caption(txt(row.get("observacao")))
    st.markdown("</div>", unsafe_allow_html=True)


def minhas_tarefas(user):
    ref = app_frame("Minhas tarefas", f"Atividades atribuídas a {user}")
    tarefas = preparar_tarefas(ref)
    minhas = tarefas[tarefas.apply(lambda r: pertence_usuario(r, user), axis=1)] if not tarefas.empty else tarefas
    abertas = minhas[~minhas.apply(lambda r: is_concluida(r, ref), axis=1)] if not minhas.empty else minhas
    if abertas.empty:
        st.success("Você não possui tarefas pendentes.")
    else:
        for _, row in abertas.iterrows():
            task_card(row, user, "minhas")


def departamento_page(user, departamento):
    ref = app_frame(departamento, f"Tarefas do departamento {departamento}")
    tarefas = preparar_tarefas(ref)
    if tarefas.empty:
        st.info("Nenhuma tarefa cadastrada.")
        return
    base = tarefas[tarefas["departamento"].astype(str).str.lower() == departamento.lower()]
    busca = st.text_input("Buscar tarefa")
    mostrar_concluidas = st.toggle("Mostrar concluídas", value=False)
    if not mostrar_concluidas:
        base = base[~base.apply(lambda r: is_concluida(r, ref), axis=1)]
    if busca:
        b = busca.lower()
        base = base[base["tarefa"].astype(str).str.lower().str.contains(b, na=False) | base["descricao"].astype(str).str.lower().str.contains(b, na=False)]
    if base.empty:
        st.info("Nenhuma tarefa encontrada.")
    else:
        for _, row in base.iterrows():
            task_card(row, user, f"dep_{departamento}")


def cadastro(user):
    ref = app_frame("Cadastro de tarefas", "Inclua novas atividades no sistema")
    st.info("Ao escolher um Projeto no cadastro, esta tarefa será vinculada automaticamente ao projeto e contará no progresso.")
    deps = listar_departamentos() or ["Contas a Receber", "Contas a Pagar", "Contabilidade", "Controladoria", "Tesouraria"]
    usuarios = listar_usuarios()
    responsaveis = [""] + (usuarios["nome"].tolist() if not usuarios.empty else [])
    projetos = listar_projetos()
    projetos_l = [""] + (projetos["projeto"].dropna().tolist() if not projetos.empty else [])
    with st.form("nova_tarefa"):
        c1, c2 = st.columns(2)
        with c1:
            tarefa = st.text_input("Tarefa")
            descricao = st.text_area("Descrição")
            departamento = st.selectbox("Departamento", deps)
            projeto = st.selectbox("Projeto", projetos_l)
            responsavel = st.selectbox("Responsável", responsaveis)
        with c2:
            periodicidade = st.selectbox("Periodicidade", ["Diario", "Semanal", "Mensal", "Unica"])
            obrigatoria = st.selectbox("Obrigatória", ["Não", "Sim"])
            prioridade = st.selectbox("Prioridade", ["Normal", "Alta", "Crítica", "Baixa"])
            dependencia = st.text_input("Dependência")
            data_inicio = st.date_input("Data de início", value=hoje())
            prazo_limite = st.time_input("Prazo limite")
        observacao = st.text_area("Observação inicial")
        salvar = st.form_submit_button("Salvar tarefa")
    if salvar:
        if not tarefa:
            st.error("Informe a tarefa.")
            return
        dados = {"tarefa": tarefa, "descricao": descricao, "departamento": normalizar_departamento(departamento), "projeto": projeto, "responsavel": responsavel, "periodicidade": periodicidade, "obrigatoria": obrigatoria, "prioridade": prioridade, "dependencia": dependencia, "prazo_limite": prazo_limite.strftime("%H:%M:%S"), "data_inicio": data_inicio.strftime("%d/%m/%Y"), "observacao": observacao}
        criar_tarefa(dados, user)
        st.success("Tarefa criada.")
        st.rerun()


def editar(user):
    ref = app_frame("Editar tarefas", "Altere ou arquive atividades existentes")
    tarefas = listar_tarefas(ativas=True)
    if tarefas.empty:
        st.info("Nenhuma tarefa ativa.")
        return
    tarefas["opcao"] = tarefas["id"].astype(str) + " - " + tarefas["tarefa"].astype(str)
    opcao = st.selectbox("Selecione a tarefa", tarefas["opcao"].tolist())
    id_tarefa = int(opcao.split(" - ")[0])
    row = tarefas[tarefas["id"] == id_tarefa].iloc[0]
    deps = listar_departamentos() or ["Contas a Receber", "Contas a Pagar", "Contabilidade", "Controladoria", "Tesouraria"]
    usuarios = listar_usuarios()
    responsaveis = [""] + (usuarios["nome"].tolist() if not usuarios.empty else [])
    with st.form("editar_tarefa"):
        tarefa = st.text_input("Tarefa", value=txt(row.get("tarefa")))
        descricao = st.text_area("Descrição", value=txt(row.get("descricao")))
        departamento = st.selectbox("Departamento", deps, index=deps.index(txt(row.get("departamento"))) if txt(row.get("departamento")) in deps else 0)
        responsavel = st.selectbox("Responsável", responsaveis, index=responsaveis.index(txt(row.get("responsavel"))) if txt(row.get("responsavel")) in responsaveis else 0)
        periodicidade = st.selectbox("Periodicidade", ["Diario", "Semanal", "Mensal", "Unica"], index=0)
        prioridade = st.selectbox("Prioridade", ["Normal", "Alta", "Crítica", "Baixa"], index=0)
        status = st.selectbox("Status", ["Pendente", "Em andamento", "Concluída", "Suspensa"], index=0)
        observacao = st.text_area("Observação", value=txt(row.get("observacao")))
        salvar = st.form_submit_button("Salvar alterações")
    if salvar:
        dados = {"tarefa": tarefa, "descricao": descricao, "departamento": normalizar_departamento(departamento), "projeto": txt(row.get("projeto")), "responsavel": responsavel, "periodicidade": periodicidade, "obrigatoria": txt(row.get("obrigatoria")) or "Não", "prioridade": prioridade, "dependencia": txt(row.get("dependencia")), "prazo_limite": txt(row.get("prazo_limite")), "data_inicio": txt(row.get("data_inicio")), "status": status, "observacao": observacao}
        editar_tarefa(id_tarefa, dados, user)
        st.success("Tarefa atualizada.")
        st.rerun()
    if st.button("Arquivar tarefa"):
        arquivar_tarefa(id_tarefa, user)
        st.success("Tarefa arquivada.")
        st.rerun()


def pendencias_observacao(user):
    ref = app_frame("Pendências com observação", "Tarefas abertas que possuem justificativas")
    tarefas = preparar_tarefas(ref)
    if tarefas.empty:
        st.info("Sem tarefas.")
        return
    abertas = tarefas[~tarefas.apply(lambda r: is_concluida(r, ref), axis=1)]
    linhas = []
    for _, row in abertas.iterrows():
        obs = ultima_observacao(row.get("id"))
        if obs:
            r = row.copy()
            r["ultima_observacao"] = obs
            linhas.append(r)
    if not linhas:
        st.success("Nenhuma pendência aberta com observação.")
        return
    df = pd.DataFrame(linhas)
    st.dataframe(df[["id", "tarefa", "departamento", "responsavel", "status", "ultima_observacao"]], use_container_width=True, hide_index=True)


def arquivadas(user):
    ref = app_frame("Tarefas arquivadas", "Consulta e recuperação de tarefas")
    tarefas = listar_tarefas(ativas=False)
    arq = tarefas[tarefas["ativa"] == "Não"] if not tarefas.empty else tarefas
    if arq.empty:
        st.info("Nenhuma tarefa arquivada.")
        return
    for _, row in arq.iterrows():
        with st.container(border=True):
            st.markdown(f"### {row.get('tarefa')}")
            st.caption(f"{row.get('departamento')} | {row.get('responsavel')}")
            if st.button("Reativar", key=f"reativar_{row.get('id')}"):
                reativar_tarefa(int(row.get("id")), user)
                st.rerun()


def calendario(user):
    ref = app_frame("Calendário", "Visão por data")
    tarefas = listar_tarefas(ativas=True)
    if tarefas.empty:
        st.info("Sem tarefas.")
        return
    tarefas["classificacao_data"] = tarefas.apply(lambda r: classificar(r, ref), axis=1)
    previstas = tarefas[(tarefas.apply(lambda r: periodic_on_date(r, ref), axis=1)) & (~tarefas.apply(lambda r: is_concluida(r, ref), axis=1))]
    concl = tarefas[tarefas.apply(lambda r: is_concluida(r, ref), axis=1)]
    c1, c2 = st.columns(2)
    with c1: metric_card("Previstas", len(previstas), ref.strftime("%d/%m/%Y"), "#f59e0b")
    with c2: metric_card("Concluídas", len(concl), ref.strftime("%d/%m/%Y"), "#16a34a")
    st.subheader("Previstas")
    st.dataframe(previstas, use_container_width=True, hide_index=True)
    st.subheader("Concluídas")
    st.dataframe(concl, use_container_width=True, hide_index=True)



def calcular_status_projeto(nome_projeto):
    tarefas = listar_tarefas(ativas=True)

    if tarefas.empty or not txt(nome_projeto):
        return {
            "total": 0,
            "concluidas": 0,
            "pendentes": 0,
            "em_andamento": 0,
            "atrasadas": 0,
            "percentual": 0,
            "status_calc": "Sem tarefas",
        }

    vinculadas = tarefas[tarefas["projeto"].astype(str).str.strip().str.lower() == txt(nome_projeto).lower()].copy()

    if vinculadas.empty:
        return {
            "total": 0,
            "concluidas": 0,
            "pendentes": 0,
            "em_andamento": 0,
            "atrasadas": 0,
            "percentual": 0,
            "status_calc": "Sem tarefas",
        }

    vinculadas["classificacao"] = vinculadas.apply(lambda r: classificar(r, hoje()), axis=1)

    total = len(vinculadas)
    concluidas = len(vinculadas[vinculadas.apply(lambda r: is_concluida(r, hoje()), axis=1)])
    pendentes = len(vinculadas[~vinculadas.apply(lambda r: is_concluida(r, hoje()), axis=1)])
    em_andamento = len(vinculadas[vinculadas["status"].astype(str).str.lower() == "em andamento"])
    atrasadas = len(vinculadas[vinculadas.apply(lambda r: eh_pendencia_anterior(r, data_referencia_global()), axis=1)])

    percentual = round((concluidas / total) * 100, 1) if total else 0

    if percentual >= 100:
        status_calc = "Concluído"
    elif atrasadas > 0:
        status_calc = "Em risco"
    elif em_andamento > 0:
        status_calc = "Em andamento"
    else:
        status_calc = "Aberto"

    return {
        "total": total,
        "concluidas": concluidas,
        "pendentes": pendentes,
        "em_andamento": em_andamento,
        "atrasadas": atrasadas,
        "percentual": percentual,
        "status_calc": status_calc,
    }


def projeto_card(row, user, prefix="proj"):
    nome = txt(row.get("projeto"))
    objetivo = txt(row.get("objetivo"))
    departamento = txt(row.get("departamento"))
    responsavel = txt(row.get("responsavel"))
    prazo = txt(row.get("prazo_final"))
    proxima = txt(row.get("proxima_etapa"))
    observacao = txt(row.get("observacao"))

    stats = calcular_status_projeto(nome)
    pct = float(stats["percentual"])

    cor = "#16a34a" if pct >= 100 else "#dc2626" if stats["atrasadas"] > 0 else "#f59e0b" if pct > 0 else "#2563eb"

    with st.container(border=True):
        c1, c2 = st.columns([4, 1.2])

        with c1:
            st.markdown(f"### 📁 {nome}")
            if objetivo:
                st.caption(objetivo)

            tags = ""
            tags += f"<span class='tag tag-blue'>{departamento or '-'}</span>"
            tags += f"<span class='tag tag-purple'>{stats['status_calc']}</span>"
            if responsavel:
                tags += f"<span class='tag tag-green'>Responsável: {responsavel}</span>"
            if prazo:
                tags += f"<span class='tag tag-yellow'>Prazo: {prazo}</span>"
            st.markdown(tags, unsafe_allow_html=True)

            st.progress(min(max(pct / 100, 0), 1))
            st.caption(
                f"{stats['concluidas']} concluídas | {stats['pendentes']} pendentes | "
                f"{stats['em_andamento']} em andamento | {stats['atrasadas']} atrasadas | "
                f"{stats['total']} tarefas vinculadas"
            )

            if proxima:
                st.info(f"Próxima etapa: {proxima}")
            if observacao:
                st.caption(f"Observação: {observacao}")

        with c2:
            metric_card("Progresso", f"{pct:.0f}%", stats["status_calc"], cor)

        with st.expander("Ver tarefas vinculadas", expanded=False):
            tarefas = listar_tarefas(ativas=True)
            if tarefas.empty:
                st.info("Nenhuma tarefa cadastrada.")
            else:
                vinculadas = tarefas[tarefas["projeto"].astype(str).str.strip().str.lower() == nome.lower()].copy()
                if vinculadas.empty:
                    st.info("Nenhuma tarefa vinculada a este projeto.")
                else:
                    for _, tarefa in vinculadas.iterrows():
                        task_card(tarefa, user, f"{prefix}_{int(row.get('id'))}")

        with st.expander("Comentário do projeto", expanded=False):
            comentario = st.text_area("Registrar comentário do projeto", key=f"coment_proj_{prefix}_{int(row.get('id'))}")
            if st.button("Salvar comentário", key=f"save_coment_proj_{prefix}_{int(row.get('id'))}"):
                if comentario.strip():
                    # usa histórico com id_tarefa vazio para registrar eventos do projeto
                    from database import registrar_historico
                    registrar_historico(None, f"Projeto: {nome}", user, "Comentário Projeto", comentario)
                    st.success("Comentário registrado no histórico.")
                    st.rerun()
                else:
                    st.warning("Digite um comentário.")


def cadastrar_projeto_sqlite(dados):
    from database import executar, agora_br
    executar(
        """
        INSERT INTO projetos (
            projeto, objetivo, departamento, responsavel, data_inicio, prazo_final,
            status, percentual, proxima_etapa, observacao, ativo
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Sim')
        """,
        (
            dados.get("projeto", ""),
            dados.get("objetivo", ""),
            dados.get("departamento", ""),
            dados.get("responsavel", ""),
            dados.get("data_inicio", ""),
            dados.get("prazo_final", ""),
            dados.get("status", "Em andamento"),
            0,
            dados.get("proxima_etapa", ""),
            dados.get("observacao", ""),
        )
    )


def projetos_resumo_df():
    projetos_df = listar_projetos()
    if projetos_df.empty:
        return pd.DataFrame()

    linhas = []
    for _, row in projetos_df.iterrows():
        nome = txt(row.get("projeto"))
        stats = calcular_status_projeto(nome)
        linhas.append({
            "Projeto": nome,
            "Departamento": txt(row.get("departamento")),
            "Responsável": txt(row.get("responsavel")),
            "Prazo": txt(row.get("prazo_final")),
            "Status": stats["status_calc"],
            "Progresso %": stats["percentual"],
            "Tarefas": stats["total"],
            "Concluídas": stats["concluidas"],
            "Pendentes": stats["pendentes"],
            "Em andamento": stats["em_andamento"],
            "Atrasadas": stats["atrasadas"],
            "Próxima etapa": txt(row.get("proxima_etapa")),
        })
    return pd.DataFrame(linhas)



def projetos(user):
    ref = app_frame("Projetos", "Gerenciamento integrado dos projetos e tarefas vinculadas")

    projetos_df = listar_projetos()

    with st.expander("➕ Cadastrar novo projeto", expanded=False):
        deps = listar_departamentos() or ["Contas a Receber", "Contas a Pagar", "Contabilidade", "Controladoria", "Tesouraria"]
        usuarios = listar_usuarios()
        responsaveis = [""] + (usuarios["nome"].tolist() if not usuarios.empty else [])

        with st.form("novo_projeto"):
            c1, c2 = st.columns(2)
            with c1:
                projeto = st.text_input("Nome do projeto")
                objetivo = st.text_area("Objetivo")
                departamento = st.selectbox("Departamento", deps)
                responsavel = st.selectbox("Responsável", responsaveis)
            with c2:
                data_inicio = st.date_input("Data início", value=hoje(), key="proj_data_inicio")
                prazo_final = st.date_input("Prazo final", value=hoje(), key="proj_prazo_final")
                proxima_etapa = st.text_input("Próxima etapa")
                status = st.selectbox("Status", ["Em andamento", "Aberto", "Suspenso", "Concluído"])
            observacao = st.text_area("Observação")
            salvar = st.form_submit_button("Salvar projeto")

        if salvar:
            if not projeto:
                st.error("Informe o nome do projeto.")
            else:
                cadastrar_projeto_sqlite({
                    "projeto": projeto,
                    "objetivo": objetivo,
                    "departamento": normalizar_departamento(departamento),
                    "responsavel": responsavel,
                    "data_inicio": data_inicio.strftime("%d/%m/%Y"),
                    "prazo_final": prazo_final.strftime("%d/%m/%Y"),
                    "status": status,
                    "proxima_etapa": proxima_etapa,
                    "observacao": observacao,
                })
                st.success("Projeto cadastrado.")
                st.rerun()

    resumo = projetos_resumo_df()

    if resumo.empty:
        st.info("Nenhum projeto cadastrado.")
        return

    c1, c2, c3, c4 = st.columns(4)
    total = len(resumo)
    concluidos = len(resumo[resumo["Status"] == "Concluído"])
    risco = len(resumo[resumo["Status"] == "Em risco"])
    andamento = len(resumo[resumo["Status"].isin(["Em andamento", "Aberto"])])

    with c1:
        metric_card("Projetos", total, "Total", "#2563eb")
    with c2:
        metric_card("Em andamento", andamento, "Abertos", "#f59e0b")
    with c3:
        metric_card("Em risco", risco, "Com atrasos", "#dc2626")
    with c4:
        metric_card("Concluídos", concluidos, "100% tarefas", "#16a34a")

    st.markdown("<div class='panel'><div class='panel-title'>📊 Resumo dos projetos</div>", unsafe_allow_html=True)
    st.dataframe(resumo, use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='panel'><div class='panel-title'>📁 Projetos</div>", unsafe_allow_html=True)
    for _, row in projetos_df.iterrows():
        projeto_card(row, user, "projetos")
    st.markdown("</div>", unsafe_allow_html=True)



def historico():
    ref = app_frame("Histórico", "Registro das movimentações")
    df = listar_historico()
    if df.empty:
        st.info("Sem histórico.")
        return

    mostrar_tudo = st.toggle("Mostrar histórico completo", value=False)

    if not mostrar_tudo and "data" in df.columns:
        data_ref_str = ref.strftime("%d/%m/%Y")
        df = df[df["data"].astype(str) == data_ref_str]

    if df.empty:
        st.info("Nenhuma movimentação para a data selecionada.")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)


def exportar():
    ref = app_frame("Exportar Excel", "Baixe uma cópia completa da base SQLite")
    saida = Path("Agenda_Exportada.xlsx")
    if st.button("Gerar Excel"):
        exportar_excel(saida)
        st.success("Excel gerado.")
    if saida.exists():
        st.download_button("Baixar Agenda_Exportada.xlsx", data=saida.read_bytes(), file_name="Agenda_Exportada.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def admin_sqlite():
    ref = app_frame("Admin SQLite", "Ferramentas de suporte da base")
    tarefas = listar_tarefas(ativas=False)
    usuarios = listar_usuarios()
    projetos_df = listar_projetos()
    st.write(f"Tarefas no banco: **{len(tarefas)}**")
    st.write(f"Usuários no banco: **{len(usuarios)}**")
    st.write(f"Projetos no banco: **{len(projetos_df)}**")
    st.divider()
    st.subheader("Diagnóstico de periodicidade")
    if not tarefas.empty and "periodicidade" in tarefas.columns:
        diag = tarefas["periodicidade"].fillna("").astype(str).value_counts().reset_index()
        diag.columns = ["Periodicidade", "Quantidade"]
        st.dataframe(diag, use_container_width=True, hide_index=True)

        tarefas_diag = tarefas.copy()
        tarefas_diag["periodicidade_normalizada"] = tarefas_diag["periodicidade"].apply(normalizar_periodicidade)
        tarefas_diag["aparece_hoje"] = tarefas_diag.apply(lambda r: periodic_on_date(r, hoje()), axis=1)
        st.write(f"Tarefas que devem aparecer hoje: **{int(tarefas_diag['aparece_hoje'].sum())}**")
        st.caption("Regra de dias úteis: tarefas diárias aparecem apenas de segunda a sexta.")
        tarefas_diag["pendencia_anterior_calc"] = tarefas_diag.apply(lambda r: eh_pendencia_anterior(r, data_referencia_global()), axis=1)
        st.write(f"Pendências anteriores calculadas: **{int(tarefas_diag['pendencia_anterior_calc'].sum())}**")

    st.warning("A migração automática está bloqueada quando o banco já possui dados, para evitar duplicações.")
    if st.button("Migrar Agenda.xlsx apenas se o banco estiver vazio"):
        ok, msg = migrar_excel_para_sqlite("Agenda.xlsx", DB_PATH, somente_se_vazio=True)
        st.success(msg if ok else msg)
        st.rerun()


def main():
    user, page, deps = sidebar()
    if page == "Dashboard": dashboard(user)
    elif page == "Coordenação": coordenacao(user)
    elif page == "Minhas tarefas": minhas_tarefas(user)
    elif page in deps: departamento_page(user, page)
    elif page == "Projetos": projetos(user)
    elif page == "Calendário": calendario(user)
    elif page == "Cadastro de tarefas": cadastro(user)
    elif page == "Editar tarefas": editar(user)
    elif page == "Pendências com observação": pendencias_observacao(user)
    elif page == "Tarefas arquivadas": arquivadas(user)
    elif page == "Histórico": historico()
    elif page == "Exportar Excel": exportar()
    elif page == "Admin SQLite": admin_sqlite()


if __name__ == "__main__":
    main()
