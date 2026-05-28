
# -*- coding: utf-8 -*-
from pathlib import Path
from datetime import datetime, date, time
from zoneinfo import ZoneInfo
import hashlib
import shutil

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Agenda Operacional",
    page_icon="✅",
    layout="wide",
    initial_sidebar_state="expanded"
)

ARQUIVO_EXCEL = "Agenda.xlsx"
FUSO_HORARIO = ZoneInfo("America/Sao_Paulo")


def agora_brasilia():
    return datetime.now(FUSO_HORARIO)

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


st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #0f172a 0%, #111827 60%, #020617 100%); }
section[data-testid="stSidebar"] * { color: #f8fafc !important; }
.main-title { font-size: 34px; font-weight: 800; color: #111827; margin-bottom: 0; }
.subtitle { color: #64748b; font-size: 15px; margin-top: 2px; }
.metric-card { background: white; border: 1px solid #e5e7eb; border-radius: 22px; padding: 22px; box-shadow: 0 8px 28px rgba(15, 23, 42, 0.06); min-height: 135px; }
.metric-label { color: #334155; font-size: 15px; font-weight: 700; }
.metric-value { font-size: 36px; font-weight: 800; margin-top: 8px; }
.metric-detail { font-size: 13px; margin-top: 4px; }
.panel { background: white; border: 1px solid #e5e7eb; border-radius: 24px; padding: 22px; box-shadow: 0 8px 28px rgba(15, 23, 42, 0.05); margin-top: 18px; }
.highlight-panel { background: linear-gradient(135deg, #eef2ff 0%, #f5f3ff 100%); border: 1px solid #c7d2fe; border-radius: 24px; padding: 22px; box-shadow: 0 8px 28px rgba(79, 70, 229, 0.08); margin-top: 18px; }
.panel-title { font-size: 22px; font-weight: 800; color: #111827; margin-bottom: 18px; }
.task-title { font-size: 18px; font-weight: 800; color: #111827; }
.task-title-done { font-size: 18px; font-weight: 800; color: #64748b; text-decoration: line-through; }
.task-desc { color: #64748b; font-size: 14px; margin-top: 6px; }
.task-muted { opacity: 0.72; }
.tag { display: inline-block; padding: 6px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; margin-right: 6px; margin-top: 8px; }
.tag-blue { background:#dbeafe; color:#1d4ed8; }
.tag-green { background:#dcfce7; color:#15803d; }
.tag-red { background:#fee2e2; color:#b91c1c; }
.tag-yellow { background:#fef3c7; color:#b45309; }
.tag-purple { background:#ede9fe; color:#6d28d9; }
.stButton > button { border-radius: 14px; min-height: 44px; font-weight: 700; border: 1px solid #d1d5db; }
.stButton > button:hover { border-color: #7c3aed; color: #7c3aed; }

/* Correção de scroll */
html, body {
    overflow-y: auto !important;
}

div[data-testid="stAppViewContainer"] {
    overflow-y: auto !important;
}

section.main {
    overflow-y: auto !important;
}

section[data-testid="stSidebar"] {
    overflow-y: auto !important;
    max-height: 100vh !important;
}

div[data-testid="stSidebarContent"] {
    overflow-y: auto !important;
    max-height: 100vh !important;
    padding-bottom: 40px !important;
}

div.block-container {
    padding-top: 2rem !important;
    padding-bottom: 5rem !important;
    overflow-y: visible !important;
}

[data-testid="stVerticalBlock"] {
    overflow: visible !important;
}


section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: #ffffff !important;
    color: #111827 !important;
    border-radius: 12px !important;
    border: 1px solid #334155 !important;
}

section[data-testid="stSidebar"] div[data-baseweb="select"] span,
section[data-testid="stSidebar"] div[data-baseweb="select"] div {
    color: #111827 !important;
}

section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p {
    color: #e5e7eb !important;
}


.coord-grid-card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 22px;
    padding: 18px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
    margin-bottom: 14px;
}

.coord-title {
    font-size: 18px;
    font-weight: 800;
    color: #111827;
    margin-bottom: 8px;
}

.coord-sub {
    color: #64748b;
    font-size: 13px;
}

.sla-red { border-left: 7px solid #dc2626 !important; }
.sla-yellow { border-left: 7px solid #f59e0b !important; }
.sla-green { border-left: 7px solid #16a34a !important; }
.sla-purple { border-left: 7px solid #7c3aed !important; }
.sla-gray { border-left: 7px solid #94a3b8 !important; }

.kanban-col {
    background: #f8fafc;
    border: 1px solid #e5e7eb;
    border-radius: 22px;
    padding: 16px;
    min-height: 360px;
}

.kanban-header {
    font-size: 18px;
    font-weight: 800;
    color: #0f172a;
    margin-bottom: 12px;
}

.feed-item {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
    padding: 12px 14px;
    margin-bottom: 10px;
    color: #334155;
}


/* ================= V12 UI/UX ================= */
:root {
    --primary: #6d28d9;
    --primary2: #2563eb;
    --ink: #0f172a;
    --muted: #64748b;
    --surface: #ffffff;
    --soft: #f8fafc;
}

div.block-container {
    padding-top: 1.3rem !important;
    max-width: 1480px !important;
}

section[data-testid="stSidebar"] {
    background: radial-gradient(circle at top left, #312e81 0%, #111827 38%, #020617 100%) !important;
    border-right: 1px solid rgba(255,255,255,0.08);
}

section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #ffffff !important;
    letter-spacing: -0.03em;
}

section[data-testid="stSidebar"] div[role="radiogroup"] label {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
    border-radius: 14px !important;
    padding: 8px 10px !important;
    margin-bottom: 6px !important;
}

section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    background: rgba(124,58,237,0.28) !important;
}

section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background: #ffffff !important;
    border-radius: 14px !important;
    border: 1px solid #a78bfa !important;
    min-height: 44px;
}

section[data-testid="stSidebar"] div[data-baseweb="select"] * {
    color: #111827 !important;
}

.main-title {
    font-size: 40px !important;
    letter-spacing: -0.04em;
    background: linear-gradient(90deg, #111827, #4338ca, #7c3aed);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}

.subtitle {
    font-size: 16px !important;
}

.metric-card {
    border: 1px solid rgba(226,232,240,0.9) !important;
    border-radius: 28px !important;
    background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%) !important;
    box-shadow: 0 16px 45px rgba(15,23,42,0.08) !important;
    transition: transform .18s ease, box-shadow .18s ease;
}

.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 22px 55px rgba(15,23,42,0.11) !important;
}

.panel, .highlight-panel {
    border-radius: 30px !important;
    border: 1px solid #e2e8f0 !important;
    box-shadow: 0 18px 45px rgba(15,23,42,0.07) !important;
}

.highlight-panel {
    background: linear-gradient(135deg, #eef2ff 0%, #faf5ff 55%, #eff6ff 100%) !important;
}

div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 22px !important;
    border-color: #e2e8f0 !important;
    box-shadow: 0 8px 24px rgba(15,23,42,0.05);
}

.task-title, .task-title-done {
    letter-spacing: -0.02em;
}

.tag {
    border: 1px solid rgba(255,255,255,0.4);
}

.coord-hero {
    background: linear-gradient(135deg, #0f172a 0%, #312e81 52%, #6d28d9 100%);
    color: #ffffff;
    border-radius: 32px;
    padding: 28px 30px;
    box-shadow: 0 22px 60px rgba(49,46,129,0.26);
    margin-bottom: 20px;
}

.coord-hero h2 {
    margin: 0;
    font-size: 30px;
    letter-spacing: -0.04em;
}

.coord-hero p {
    margin: 6px 0 0 0;
    color: #ddd6fe;
}

.v12-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 24px;
    padding: 18px;
    margin-bottom: 14px;
    box-shadow: 0 12px 35px rgba(15,23,42,0.06);
}

.v12-card-title {
    font-weight: 850;
    color: #0f172a;
    font-size: 17px;
    letter-spacing: -0.02em;
}

.v12-card-meta {
    color: #64748b;
    font-size: 13px;
    margin-top: 6px;
}

.kanban-col {
    background: linear-gradient(180deg, #f8fafc 0%, #ffffff 100%);
    border: 1px solid #e2e8f0;
    border-radius: 28px;
    padding: 16px;
    min-height: 420px;
    box-shadow: inset 0 1px 0 rgba(255,255,255,0.7);
}

.kanban-header {
    font-size: 18px;
    font-weight: 850;
    color: #0f172a;
    letter-spacing: -0.03em;
    margin-bottom: 14px;
}

.feed-item {
    background: linear-gradient(90deg, #ffffff 0%, #f8fafc 100%);
    border: 1px solid #e2e8f0;
    border-radius: 18px;
    padding: 13px 15px;
    margin-bottom: 10px;
    box-shadow: 0 8px 25px rgba(15,23,42,0.04);
}

.sla-red { border-left: 8px solid #dc2626 !important; }
.sla-yellow { border-left: 8px solid #f59e0b !important; }
.sla-green { border-left: 8px solid #16a34a !important; }
.sla-purple { border-left: 8px solid #7c3aed !important; }
.sla-gray { border-left: 8px solid #94a3b8 !important; }

.stButton > button {
    border-radius: 16px !important;
}


/* ================= V13 PREMIUM VISUAL ================= */

/* Ocultar elementos padrão do Streamlit */
header[data-testid="stHeader"] { display: none !important; }
div[data-testid="stToolbar"] { display: none !important; }
div[data-testid="stDecoration"] { display: none !important; }
#MainMenu { visibility: hidden !important; }
footer { visibility: hidden !important; }
[data-testid="stStatusWidget"] { display: none !important; }

.stApp { background: #f4f7fb !important; }

div.block-container {
    padding-top: 2rem !important;
    padding-left: 3rem !important;
    padding-right: 3rem !important;
    max-width: 1600px !important;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #020617 0%, #111827 45%, #1e1b4b 100%) !important;
    width: 310px !important;
    min-width: 310px !important;
    border-right: 1px solid rgba(255,255,255,0.08);
    box-shadow: 12px 0 40px rgba(15, 23, 42, 0.18);
}

section[data-testid="stSidebar"] > div {
    padding: 2rem 1.3rem !important;
}

section[data-testid="stSidebar"] h2 {
    font-size: 24px !important;
    font-weight: 900 !important;
    line-height: 1.05 !important;
}

section[data-testid="stSidebar"] h3 {
    font-size: 18px !important;
    color: #c4b5fd !important;
}

section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%) !important;
    color: #0f172a !important;
    border-radius: 18px !important;
    border: 2px solid #8b5cf6 !important;
    min-height: 52px !important;
    box-shadow: 0 10px 30px rgba(124,58,237,0.18);
}

section[data-testid="stSidebar"] div[data-baseweb="select"] * {
    color: #0f172a !important;
    font-weight: 750 !important;
}

section[data-testid="stSidebar"] div[role="radiogroup"] label {
    background: transparent !important;
    border: 0 !important;
    border-radius: 18px !important;
    padding: 10px 12px !important;
    margin: 1px 0 !important;
    transition: all .18s ease;
}

section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    background: rgba(139, 92, 246, 0.20) !important;
    transform: translateX(2px);
}

section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {
    background: linear-gradient(135deg, rgba(124,58,237,0.75), rgba(37,99,235,0.55)) !important;
    box-shadow: 0 8px 24px rgba(124,58,237,0.22);
}

.main-title {
    font-size: 46px !important;
    font-weight: 950 !important;
    letter-spacing: -0.055em !important;
    color: #0f172a !important;
    background: none !important;
    -webkit-text-fill-color: #0f172a !important;
}

.subtitle {
    color: #64748b !important;
    font-size: 16px !important;
    margin-bottom: 22px !important;
}

.dashboard-hero {
    background: linear-gradient(135deg, #0f172a 0%, #312e81 45%, #6d28d9 100%);
    border-radius: 34px;
    padding: 30px 34px;
    color: #ffffff;
    margin-bottom: 24px;
    box-shadow: 0 24px 70px rgba(49, 46, 129, 0.30);
}

.dashboard-hero h1 {
    font-size: 34px;
    font-weight: 950;
    margin: 0;
    letter-spacing: -0.045em;
}

.dashboard-hero p {
    color: #ddd6fe;
    margin: 8px 0 0 0;
    font-size: 16px;
}

.metric-card {
    border-radius: 30px !important;
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    box-shadow: 0 18px 45px rgba(15, 23, 42, 0.08) !important;
    padding: 24px !important;
}

.metric-label {
    font-size: 14px !important;
    text-transform: uppercase;
    letter-spacing: .03em;
    color: #64748b !important;
}

.metric-value {
    font-size: 42px !important;
    font-weight: 950 !important;
}

.panel, .highlight-panel {
    border-radius: 34px !important;
    border: 1px solid #e2e8f0 !important;
    box-shadow: 0 18px 55px rgba(15, 23, 42, 0.08) !important;
    background: #ffffff !important;
}

.highlight-panel {
    background: linear-gradient(135deg, #ffffff 0%, #f5f3ff 100%) !important;
    border: 1px solid #c4b5fd !important;
}

div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 26px !important;
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    box-shadow: 0 12px 35px rgba(15,23,42,0.07) !important;
}

.task-title {
    font-size: 20px !important;
    font-weight: 900 !important;
}

.task-title-done {
    font-size: 20px !important;
    font-weight: 850 !important;
    color: #94a3b8 !important;
    text-decoration: line-through;
}

.tag {
    padding: 7px 12px !important;
    font-size: 12px !important;
    font-weight: 850 !important;
}

.stButton > button {
    background: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 18px !important;
    min-height: 48px !important;
    font-weight: 850 !important;
    color: #0f172a !important;
    box-shadow: 0 8px 22px rgba(15,23,42,0.05);
}

.stButton > button:hover {
    border-color: #7c3aed !important;
    color: #6d28d9 !important;
    box-shadow: 0 10px 28px rgba(124,58,237,0.16);
}

input, textarea { border-radius: 16px !important; }

.top-date-pill {
    display: inline-block;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 999px;
    padding: 10px 16px;
    font-weight: 850;
    color: #0f172a;
    box-shadow: 0 10px 30px rgba(15,23,42,0.06);
}


/* ================= V14 SIDEBAR FIXA ================= */

/* Mantém a sidebar sempre visível e com contraste */
section[data-testid="stSidebar"] {
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
    min-width: 310px !important;
    width: 310px !important;
    max-width: 310px !important;
    position: fixed !important;
    left: 0 !important;
    top: 0 !important;
    height: 100vh !important;
    z-index: 999999 !important;
}

/* Empurra o conteúdo para não ficar atrás da sidebar fixa */
section.main,
div[data-testid="stAppViewContainer"] main {
    margin-left: 310px !important;
}

/* Remove/oculta o botão de recolher a sidebar quando ela está aberta */
button[data-testid="stSidebarCollapseButton"],
[data-testid="stSidebarCollapseButton"] {
    display: none !important;
    visibility: hidden !important;
}

/* Se por algum motivo o Streamlit mostrar o botão de sidebar recolhida,
   deixa ele bem visível para recuperação */
button[data-testid="collapsedControl"],
[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"] {
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
    background: #7c3aed !important;
    color: #ffffff !important;
    border-radius: 999px !important;
    border: 2px solid #ffffff !important;
    box-shadow: 0 10px 30px rgba(124,58,237,0.45) !important;
    position: fixed !important;
    left: 12px !important;
    top: 12px !important;
    z-index: 1000000 !important;
}

/* Garante que textos da sidebar fiquem legíveis */
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span {
    color: #ffffff !important;
}

/* Mas mantém o nome selecionado do usuário escuro dentro do select branco */
section[data-testid="stSidebar"] div[data-baseweb="select"] span,
section[data-testid="stSidebar"] div[data-baseweb="select"] div {
    color: #0f172a !important;
}

/* Scroll interno da sidebar fixa */
section[data-testid="stSidebar"] div[data-testid="stSidebarContent"] {
    overflow-y: auto !important;
    max-height: 100vh !important;
    padding-bottom: 60px !important;
}

/* Ajuste responsivo: em telas pequenas, permite sidebar menor */
@media (max-width: 900px) {
    section[data-testid="stSidebar"] {
        min-width: 270px !important;
        width: 270px !important;
        max-width: 270px !important;
    }

    section.main,
    div[data-testid="stAppViewContainer"] main {
        margin-left: 270px !important;
    }
}


/* ================= V15 CORREÇÃO SOBREPOSIÇÃO SIDEBAR ================= */

/* Sidebar fixa com largura menor e estável */
section[data-testid="stSidebar"] {
    min-width: 240px !important;
    width: 240px !important;
    max-width: 240px !important;
    position: fixed !important;
    left: 0 !important;
    top: 0 !important;
    height: 100vh !important;
    z-index: 999999 !important;
}

/* Área principal deslocada para a direita */
div[data-testid="stAppViewContainer"] {
    margin-left: 240px !important;
    width: calc(100vw - 240px) !important;
}

/* O main do Streamlit precisa respeitar a nova área útil */
div[data-testid="stAppViewContainer"] > section {
    margin-left: 0 !important;
    width: 100% !important;
}

/* Container interno com respiro */
div.block-container {
    padding-left: 2.2rem !important;
    padding-right: 2.2rem !important;
    max-width: 1280px !important;
}

/* Evita que cards largos passem da tela */
.dashboard-hero,
.panel,
.highlight-panel {
    max-width: 100% !important;
    overflow: hidden !important;
}

/* Corrige sidebar em telas menores */
@media (max-width: 1100px) {
    section[data-testid="stSidebar"] {
        min-width: 220px !important;
        width: 220px !important;
        max-width: 220px !important;
    }

    div[data-testid="stAppViewContainer"] {
        margin-left: 220px !important;
        width: calc(100vw - 220px) !important;
    }

    div.block-container {
        padding-left: 1.2rem !important;
        padding-right: 1.2rem !important;
    }
}

/* Remove tentativas anteriores conflitantes */
section.main {
    margin-left: 0 !important;
}

</style>
""", unsafe_allow_html=True)


def txt(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def normalizar_departamento(v):
    t = txt(v)
    if not t:
        return ""
    mapa = {
        "contas a receber": "Contas a Receber",
        "contas a pagar": "Contas a Pagar",
        "contabilidade": "Contabilidade",
        "controladoria": "Controladoria",
        "tesouraria": "Tesouraria",
        "financeiro": "Financeiro",
        "projetos": "Projetos",
    }
    chave = " ".join(t.lower().split())
    return mapa.get(chave, t.strip().title())


def limpar_colunas(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated(keep="first")]
    return df.reset_index(drop=True)


def safe_text_df(df):
    """Evita erro de dtype ao salvar textos em colunas lidas como número/data."""
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].astype("object")
    return df


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


def done(row, ref_date=None):
    if ref_date is None:
        ref_date = date.today()

    status_ok = txt(row.get("Status")).lower() in ["concluída", "concluida", "feito", "finalizada"]
    if not status_ok:
        return False

    periodicidade = txt(row.get("Periodicidade")).lower()
    data_conclusao = parse_date(row.get("Data Conclusão"))

    if periodicidade in ["unica", "única", "pontual"]:
        return True

    return data_conclusao == ref_date


def next_id(df, col):
    if df.empty or col not in df.columns:
        return 1
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    return int(s.max()) + 1 if not s.empty else 1


def ensure_cols(df, cols):
    df = limpar_colunas(df)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols].reset_index(drop=True)
    return safe_text_df(df)


def unique_key(*parts):
    base = "|".join([str(p) for p in parts])
    return hashlib.md5(base.encode("utf-8")).hexdigest()[:12]


def classify(row, ref_date=None):
    if ref_date is None:
        ref_date = date.today()

    if done(row, ref_date):
        if ref_date == date.today():
            return "Concluída hoje"
        return "Concluída na data"

    if not is_active(row.get("Ativa", "Sim")):
        return "Arquivada"

    data_ini = parse_date(row.get("Data de Inicio"))
    periodicidade = txt(row.get("Periodicidade")).lower()

    if data_ini and data_ini > ref_date:
        return "Futura"

    if data_ini and data_ini < ref_date and periodicidade in ["unica", "única", "pontual", ""]:
        return "Pendente anterior"

    if periodic_on_date(row, ref_date):
        return "Hoje" if ref_date == date.today() else "Prevista"

    return "Futura"


def periodic_on_date(row, ref_date=None):
    if ref_date is None:
        ref_date = date.today()

    if done(row, ref_date) or not is_active(row.get("Ativa", "Sim")):
        return False

    data_ini = parse_date(row.get("Data de Inicio"))
    per = txt(row.get("Periodicidade")).lower()

    if data_ini and data_ini > ref_date:
        return False

    if per in ["diario", "diária", "diaria", "diário", "todo dia", ""]:
        return True

    if per in ["semanal", "semana"]:
        return True if not data_ini else (ref_date - data_ini).days % 7 == 0

    if per in ["mensal", "mês", "mes"]:
        return True if not data_ini else ref_date.day == data_ini.day

    if per in ["unica", "única", "pontual"]:
        return data_ini == ref_date

    return True


def periodic_today(row):
    return periodic_on_date(row, date.today())


def belongs_to_user(row, user):
    resp = txt(row.get("Responsavel")).lower()
    u = txt(user).lower()
    if not resp or not u:
        return False
    partes = resp.replace("/", ",").replace(";", ",").replace("|", ",").split(",")
    nomes = [x.strip().lower() for x in partes if x.strip()]
    return u in nomes or u in resp


def get_departamentos(data):
    tarefas = data.get("Tarefas", pd.DataFrame())
    if tarefas.empty or "Departamento" not in tarefas.columns:
        return []
    return sorted(list(set([normalizar_departamento(x) for x in tarefas["Departamento"].dropna().unique() if txt(x)])))


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
            df = pd.read_excel(path, sheet_name=aba, dtype=object)
        else:
            df = pd.DataFrame(columns=cols)
        data[aba] = ensure_cols(df, cols)

    return data


def save_data(path, data):
    path = Path(path)
    criar_backup_automatico(path)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        for aba, cols in ABAS.items():
            df = ensure_cols(data.get(aba, pd.DataFrame(columns=cols)), cols)
            df = df.fillna("")
            df.to_excel(writer, sheet_name=aba, index=False)
    st.cache_data.clear()


def get_path():
    if "excel_path" not in st.session_state:
        st.session_state["excel_path"] = ARQUIVO_EXCEL
    return st.session_state["excel_path"]


def criar_backup_automatico(path):
    try:
        origem = Path(path)
        if not origem.exists():
            return

        pasta_backup = origem.parent / "backups"
        pasta_backup.mkdir(exist_ok=True)

        agora = agora_brasilia().strftime("%Y%m%d_%H%M%S")
        destino = pasta_backup / f"Agenda_Backup_{agora}.xlsx"
        shutil.copy2(origem, destino)

        backups = sorted(pasta_backup.glob("Agenda_Backup_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
        for antigo in backups[30:]:
            try:
                antigo.unlink()
            except Exception:
                pass
    except Exception:
        pass


def atualizar_status_tarefa(data, id_tarefa, novo_status, usuario=""):
    tarefas = ensure_cols(data["Tarefas"], ABAS["Tarefas"])
    tarefas = tarefas.astype("object")

    ids = tarefas["ID"].astype(str).str.strip()
    alvo = ids == str(id_tarefa).strip()

    if not alvo.any():
        return False

    idx = tarefas.index[alvo][0]

    # Usar .at com dataframe em object evita conflito de dtype no pandas 3.
    tarefas.at[idx, "Status"] = str(novo_status)

    if novo_status == "Concluída":
        tarefas.at[idx, "Concluído Por"] = str(usuario)
        tarefas.at[idx, "Data Conclusão"] = agora_brasilia().strftime("%d/%m/%Y %H:%M:%S")
    else:
        tarefas.at[idx, "Concluído Por"] = ""
        tarefas.at[idx, "Data Conclusão"] = ""

    data["Tarefas"] = tarefas
    return True


def registrar_historico(data, id_tarefa, tarefa, usuario, status="Concluída", observacao=""):
    hist = ensure_cols(data["Historico"], ABAS["Historico"])
    hist = hist.astype("object")
    now = agora_brasilia()

    novo = {
        "ID Histórico": str(next_id(hist, "ID Histórico")),
        "ID Tarefa": str(id_tarefa),
        "Tarefa": str(tarefa),
        "Usuário": str(usuario),
        "Data": now.strftime("%d/%m/%Y"),
        "Hora": now.strftime("%H:%M:%S"),
        "Status": str(status),
        "Observação": str(observacao),
    }

    data["Historico"] = pd.concat([hist, pd.DataFrame([novo], dtype=object)], ignore_index=True)
    return data




def sla_status(row, ref_date=None):
    if ref_date is None:
        ref_date = date.today()

    if done(row, ref_date):
        return "Concluída", "sla-gray", "Concluída"

    prioridade = txt(row.get("Prioridade")).lower()
    if prioridade in ["crítica", "critica"]:
        return "Crítica", "sla-purple", "Prioridade crítica"

    data_ini = parse_date(row.get("Data de Inicio"))
    if data_ini and data_ini < ref_date:
        dias = (ref_date - data_ini).days
        return "Atrasada", "sla-red", f"Atrasada há {dias} dia(s)"

    if periodic_on_date(row, ref_date):
        return "Hoje", "sla-yellow", "Prevista para hoje"

    return "No prazo", "sla-green", "Dentro do prazo"


def tarefas_sem_responsavel(tarefas):
    return tarefas[tarefas["Responsavel"].astype(str).str.strip().isin(["", "nan", "None"])]


def ranking_operacional(data):
    tarefas = prepared_tasks(data)
    hist = ensure_cols(data["Historico"], ABAS["Historico"])

    responsaveis = sorted([x for x in tarefas["Responsavel"].dropna().astype(str).str.strip().unique() if x])
    linhas = []

    hoje_str = date.today().strftime("%d/%m/%Y")

    for resp in responsaveis:
        base = tarefas[tarefas["Responsavel"].astype(str).str.lower().str.contains(resp.lower(), na=False)]
        abertas = base[~base.apply(lambda r: done(r, date.today()), axis=1)]
        atrasadas = base[base.apply(lambda r: sla_status(r, date.today())[0] == "Atrasada", axis=1)]

        concluidas_hoje = hist[
            (hist["Usuário"].astype(str).str.lower() == resp.lower())
            & (hist["Status"].astype(str).str.lower().isin(["concluída", "concluida"]))
            & (hist["Data"].astype(str) == hoje_str)
        ]

        score = 100 - (len(atrasadas) * 12) - (len(abertas) * 3) + (len(concluidas_hoje) * 2)
        score = max(0, min(100, score))

        linhas.append({
            "Usuário": resp,
            "Pendentes": len(abertas),
            "Atrasadas": len(atrasadas),
            "Concluídas Hoje": len(concluidas_hoje),
            "Score": score
        })

    return pd.DataFrame(linhas)


def feed_operacional(data, limite=12):
    hist = ensure_cols(data["Historico"], ABAS["Historico"]).copy()
    if hist.empty:
        return hist

    hist["Ordem"] = pd.to_numeric(hist["ID Histórico"], errors="coerce")
    hist = hist.sort_values("Ordem", ascending=False).head(limite)
    return hist


def card_tarefa_resumo(row, user, prefix="coord"):
    titulo = txt(row.get("Tarefa"))
    resp = txt(row.get("Responsavel")) or "Sem responsável"
    depto = txt(row.get("Departamento"))
    prio = txt(row.get("Prioridade")) or "Normal"
    status_sla, classe_sla, detalhe = sla_status(row, date.today())
    minha = belongs_to_user(row, user)

    st.markdown(
        f"""
        <div class="coord-grid-card {classe_sla}">
            <div class="coord-title">{'⭐ ' if minha else ''}{titulo}</div>
            <div class="coord-sub">👤 {resp} &nbsp; | &nbsp; 📂 {depto or '-'} &nbsp; | &nbsp; ⚑ {prio}</div>
            <span class="tag tag-purple">{status_sla}</span>
            <span class="tag tag-gray">{detalhe}</span>
        </div>
        """,
        unsafe_allow_html=True
    )


def voltar_topo():
    st.markdown(
        """
        <div style="text-align:right; margin-top:30px;">
            <a href="#agenda-operacional" style="text-decoration:none; font-weight:700; color:#7c3aed;">
                ↑ Voltar ao topo
            </a>
        </div>
        """,
        unsafe_allow_html=True
    )


def header(title, subtitle):
    col1, col2 = st.columns([4, 1.5])

    with col1:
        st.markdown(f"<div id='agenda-operacional' class='main-title'>{title}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='subtitle'>{subtitle}</div>", unsafe_allow_html=True)

    with col2:
        hoje = agora_brasilia()
        st.markdown(
            f"""
            <div style='text-align:right; margin-top:8px;'>
                <span class="top-date-pill">📅 {hoje.strftime('%d/%m/%Y')} • {hoje.strftime('%H:%M')}</span>
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
    if not usuarios.empty and "Nome" in usuarios.columns:
        nomes = usuarios[usuarios["Ativo"].apply(is_active)]["Nome"].dropna().astype(str).str.strip().tolist()
        nomes = sorted([n for n in nomes if n])
    else:
        nomes = []

    if not nomes:
        nomes = ["Paula"]

    user = st.sidebar.selectbox("Usuário logado", nomes)

    st.sidebar.divider()

    departamentos = get_departamentos(data)
    menu = ["Dashboard", "Coordenação", "Minhas tarefas"] + departamentos + [
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

    st.sidebar.caption("Versão 2.5.0 Sidebar ajustada")
    return user, page, departamentos


def prepared_tasks(data):
    tarefas = ensure_cols(data["Tarefas"], ABAS["Tarefas"])
    tarefas = tarefas[tarefas["Ativa"].apply(is_active)].reset_index(drop=True)
    tarefas["Departamento"] = tarefas["Departamento"].apply(normalizar_departamento)
    tarefas["Classificacao"] = tarefas.apply(lambda r: classify(r, date.today()), axis=1)
    return tarefas


def task_card(row, data, user, prefix):
    idt = txt(row.get("ID"))
    titulo = txt(row.get("Tarefa"))
    desc = txt(row.get("Descrição"))
    depto = txt(row.get("Departamento"))
    resp = txt(row.get("Responsavel"))
    prioridade = txt(row.get("Prioridade")) or "Normal"
    projeto = txt(row.get("Projeto"))
    periodicidade = txt(row.get("Periodicidade"))
    status = classify(row, date.today())
    minha = belongs_to_user(row, user)

    tag_status = "tag-green" if "Concluída" in status else "tag-red" if "anterior" in status else "tag-yellow"

    with st.container(border=True):
        c1, c2 = st.columns([4.8, 1.25])

        with c1:
            classe_titulo = "task-title-done" if done(row) else "task-title"
            st.markdown(f"<div class='{classe_titulo}'>{'⭐ ' if minha else ''}{titulo}</div>", unsafe_allow_html=True)

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
            chave_base = unique_key(prefix, idt, titulo, resp, depto)

            if done(row, date.today()):
                st.success("✅ Concluída")
                if st.button("Reabrir", key=f"reopen_{chave_base}"):
                    ok = atualizar_status_tarefa(data, idt, "Pendente", user)
                    if ok:
                        registrar_historico(data, idt, titulo, user, "Reaberta")
                        save_data(get_path(), data)
                        st.rerun()
                    else:
                        st.error("Tarefa não localizada no Excel.")
            else:
                if st.button("✅ Concluir", key=f"done_{chave_base}"):
                    ok = atualizar_status_tarefa(data, idt, "Concluída", user)
                    if ok:
                        registrar_historico(data, idt, titulo, user, "Concluída")
                        save_data(get_path(), data)
                        st.success("Tarefa concluída.")
                        st.rerun()
                    else:
                        st.error("Tarefa não localizada no Excel.")




def coordenacao_page(data, user):
    header("Coordenação", "Central gerencial da operação")

    st.markdown(
        """
        <div class="coord-hero">
            <h2>Central de Coordenação</h2>
            <p>Visão executiva de pendências, gargalos, produtividade e movimentações do time.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    tarefas = prepared_tasks(data)
    abertas = tarefas[~tarefas.apply(lambda r: done(r, date.today()), axis=1)]
    atrasadas = tarefas[tarefas.apply(lambda r: sla_status(r, date.today())[0] == "Atrasada", axis=1)]
    criticas = abertas[abertas["Prioridade"].astype(str).str.lower().isin(["crítica", "critica"])]
    sem_resp = tarefas_sem_responsavel(abertas)
    concluidas_hoje = tarefas[tarefas["Classificacao"] == "Concluída hoje"]

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        metric_card("🔴 Atrasadas", len(atrasadas), "Gargalos", "#dc2626")
    with c2:
        metric_card("🟣 Críticas", len(criticas), "Alta prioridade", "#7c3aed")
    with c3:
        metric_card("🟡 Abertas", len(abertas), "Em aberto", "#f59e0b")
    with c4:
        metric_card("✅ Concluídas hoje", len(concluidas_hoje), "Produtividade", "#16a34a")
    with c5:
        metric_card("👤 Sem responsável", len(sem_resp), "Corrigir cadastro", "#2563eb")

    st.markdown("<div class='panel'><div class='panel-title'>🧭 Kanban operacional</div>", unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)

    with k1:
        st.markdown("<div class='kanban-col'><div class='kanban-header'>🔴 Atrasadas</div>", unsafe_allow_html=True)
        if atrasadas.empty:
            st.success("Sem atrasadas.")
        else:
            for _, row in atrasadas.head(8).iterrows():
                card_tarefa_resumo(row, user, "kanban_atrasadas")
        st.markdown("</div>", unsafe_allow_html=True)

    with k2:
        st.markdown("<div class='kanban-col'><div class='kanban-header'>🟣 Críticas</div>", unsafe_allow_html=True)
        if criticas.empty:
            st.info("Sem críticas abertas.")
        else:
            for _, row in criticas.head(8).iterrows():
                card_tarefa_resumo(row, user, "kanban_criticas")
        st.markdown("</div>", unsafe_allow_html=True)

    with k3:
        st.markdown("<div class='kanban-col'><div class='kanban-header'>🟡 Hoje</div>", unsafe_allow_html=True)
        hoje = abertas[abertas.apply(lambda r: periodic_today(r), axis=1)].head(8)
        if hoje.empty:
            st.info("Sem tarefas de hoje.")
        else:
            for _, row in hoje.iterrows():
                card_tarefa_resumo(row, user, "kanban_hoje")
        st.markdown("</div>", unsafe_allow_html=True)

    with k4:
        st.markdown("<div class='kanban-col'><div class='kanban-header'>✅ Concluídas hoje</div>", unsafe_allow_html=True)
        if concluidas_hoje.empty:
            st.info("Nada concluído hoje ainda.")
        else:
            for _, row in concluidas_hoje.head(8).iterrows():
                card_tarefa_resumo(row, user, "kanban_done")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

    col_a, col_b = st.columns([1.1, 1])

    with col_a:
        st.markdown("<div class='panel'><div class='panel-title'>📊 Ranking operacional</div>", unsafe_allow_html=True)
        ranking = ranking_operacional(data)
        if ranking.empty:
            st.info("Ainda não há responsáveis suficientes para ranking.")
        else:
            st.dataframe(
                ranking.sort_values(["Atrasadas", "Pendentes", "Score"], ascending=[True, True, False]),
                use_container_width=True,
                hide_index=True
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with col_b:
        st.markdown("<div class='panel'><div class='panel-title'>🚦 Gargalos rápidos</div>", unsafe_allow_html=True)
        if atrasadas.empty and sem_resp.empty:
            st.success("Nenhum gargalo relevante agora.")
        else:
            if not atrasadas.empty:
                st.markdown("##### 🔴 Atrasadas")
                for _, row in atrasadas.head(4).iterrows():
                    card_tarefa_resumo(row, user, "gargalo_atrasadas")
            if not sem_resp.empty:
                st.markdown("##### 👤 Sem responsável")
                for _, row in sem_resp.head(4).iterrows():
                    card_tarefa_resumo(row, user, "gargalo_semresp")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='panel'><div class='panel-title'>🕘 Feed operacional</div>", unsafe_allow_html=True)
    feed = feed_operacional(data)
    if feed.empty:
        st.info("Ainda não há movimentações registradas.")
    else:
        for _, row in feed.iterrows():
            st.markdown(
                f"""
                <div class="feed-item">
                    <b>{txt(row.get('Hora'))}</b> — {txt(row.get('Usuário'))} registrou
                    <b>{txt(row.get('Status'))}</b> em “{txt(row.get('Tarefa'))}”
                    <br><span style="color:#64748b;font-size:13px;">{txt(row.get('Data'))} {txt(row.get('Observação'))}</span>
                </div>
                """,
                unsafe_allow_html=True
            )
    st.markdown("</div>", unsafe_allow_html=True)


def dashboard(data, user):
    header("Dashboard", "Visão geral das atividades")
    st.markdown(
        f"""
        <div class="dashboard-hero">
            <h1>Agenda operacional da equipe</h1>
            <p>Olá, {user}. Acompanhe prioridades, pendências, tarefas do dia e produtividade em tempo real.</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    tarefas = prepared_tasks(data)

    minhas = tarefas[tarefas.apply(lambda r: belongs_to_user(r, user), axis=1)]
    minhas_abertas = minhas[~minhas.apply(lambda r: done(r, date.today()), axis=1)]
    pend = tarefas[tarefas["Classificacao"] == "Pendente anterior"]
    hoje = tarefas[(tarefas.apply(periodic_today, axis=1)) & (~tarefas.apply(lambda r: done(r, date.today()), axis=1))]
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

    f1, f2, f3, f4 = st.columns([2, 1, 1, 1])
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
        lista = filtradas[(filtradas.apply(periodic_today, axis=1)) & (~filtradas.apply(lambda r: done(r, date.today()), axis=1))]
        if lista.empty:
            st.info("Nenhuma tarefa pendente para hoje.")
        else:
            for _, row in lista.iterrows():
                task_card(row, data, user, "hoje")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='panel'><div class='panel-title'>✅ Concluídas hoje</div>", unsafe_allow_html=True)
    concluidas_visiveis = filtradas[filtradas["Classificacao"] == "Concluída hoje"]
    if concluidas_visiveis.empty:
        st.info("Nenhuma tarefa concluída hoje.")
    else:
        for _, row in concluidas_visiveis.iterrows():
            task_card(row, data, user, "concluidas_hoje")
    st.markdown("</div>", unsafe_allow_html=True)


def minhas_tarefas_page(data, user):
    header("Minhas tarefas", f"Atividades atribuídas a {user}")
    tarefas = prepared_tasks(data)
    minhas = tarefas[tarefas.apply(lambda r: belongs_to_user(r, user), axis=1)]
    abertas = minhas[~minhas.apply(lambda r: done(r, date.today()), axis=1)]
    concluidas = minhas[minhas.apply(lambda r: done(r, date.today()), axis=1)]

    c1, c2, c3 = st.columns(3)
    with c1: metric_card("Minhas abertas", len(abertas), "Pendentes", "#7c3aed")
    with c2: metric_card("Minhas concluídas", len(concluidas), "Histórico visual", "#16a34a")
    with c3: metric_card("Total atribuídas", len(minhas), "Todas", "#2563eb")

    st.markdown("<div class='highlight-panel'><div class='panel-title'>⭐ Prioridade do usuário logado</div>", unsafe_allow_html=True)
    if abertas.empty:
        st.success("Você não possui tarefas pendentes.")
    else:
        for _, row in abertas.iterrows():
            task_card(row, data, user, "minhas_page")
    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander("Ver minhas tarefas concluídas"):
        if concluidas.empty:
            st.info("Nenhuma tarefa concluída.")
        else:
            for _, row in concluidas.iterrows():
                task_card(row, data, user, "minhas_concluidas")


def departamento_page(data, user, depto):
    header(depto, f"Tarefas do departamento {depto}")
    tarefas = prepared_tasks(data)
    tarefas = tarefas[tarefas["Departamento"].astype(str).str.lower().str.strip() == depto.lower()]

    busca = st.text_input("🔎 Buscar tarefa")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        somente_minhas = st.toggle("Mostrar somente minhas tarefas neste departamento", value=False)
    with col_f2:
        mostrar_concluidas = st.toggle("Mostrar concluídas", value=False)

    if somente_minhas:
        tarefas = tarefas[tarefas.apply(lambda r: belongs_to_user(r, user), axis=1)]

    if not mostrar_concluidas:
        tarefas = tarefas[~tarefas.apply(lambda r: done(r, date.today()), axis=1)]

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
            c1, c2 = st.columns([4, 1])
            with c1:
                st.markdown(f"### {txt(row.get('Projeto'))}")
                st.write(txt(row.get("Objetivo")))
                st.caption(
                    f"Responsável: {txt(row.get('Responsavel')) or '-'} | "
                    f"Prazo: {txt(row.get('Prazo Final')) or '-'} | "
                    f"Status: {txt(row.get('Status')) or '-'}"
                )
                st.progress(min(max(pct / 100, 0), 1))
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
            prazo = st.time_input("Prazo limite", value=time(18, 0))

        obs = st.text_area("Observação")
        ok = st.form_submit_button("Salvar tarefa")

    if ok:
        if not tarefa:
            st.error("Informe a tarefa.")
        else:
            novo = {
                "ID": str(next_id(tarefas, "ID")),
                "Tarefa": str(tarefa),
                "Descrição": str(desc),
                "Departamento": normalizar_departamento(depto),
                "Projeto": str(projeto),
                "Responsavel": str(resp),
                "Periodicidade": str(per),
                "Obrigatoria": str(obrig),
                "Prioridade": str(prio),
                "Dependencia": str(dep),
                "Prazo Limite": prazo.strftime("%H:%M:%S"),
                "Data de Inicio": data_ini.strftime("%d/%m/%Y"),
                "Status": "Pendente",
                "Concluído Por": "",
                "Data Conclusão": "",
                "Observação": str(obs),
                "Ativa": "Sim",
            }
            data["Tarefas"] = pd.concat([ensure_cols(tarefas, ABAS["Tarefas"]), pd.DataFrame([novo], dtype=object)], ignore_index=True)
            save_data(get_path(), data)
            st.success("Tarefa cadastrada.")
            st.rerun()



def editar_tarefas_page(data, user):
    header("Editar tarefas", "Altere atividades já cadastradas")

    tarefas = ensure_cols(data["Tarefas"], ABAS["Tarefas"])
    tarefas_ativas = tarefas[tarefas["Ativa"].apply(is_active)].copy()

    if tarefas_ativas.empty:
        st.info("Nenhuma tarefa ativa cadastrada.")
        return

    tarefas_ativas["Opcao"] = (
        tarefas_ativas["ID"].astype(str).str.strip()
        + " - "
        + tarefas_ativas["Tarefa"].astype(str).str.strip()
        + " | "
        + tarefas_ativas["Departamento"].astype(str).str.strip()
    )

    busca = st.text_input("🔎 Buscar tarefa para editar")

    lista_opcoes = tarefas_ativas.copy()
    if busca:
        b = busca.lower()
        lista_opcoes = lista_opcoes[
            lista_opcoes["Tarefa"].astype(str).str.lower().str.contains(b, na=False)
            | lista_opcoes["Departamento"].astype(str).str.lower().str.contains(b, na=False)
            | lista_opcoes["Responsavel"].astype(str).str.lower().str.contains(b, na=False)
        ]

    if lista_opcoes.empty:
        st.warning("Nenhuma tarefa encontrada com esse filtro.")
        return

    opcao = st.selectbox("Selecione a tarefa", lista_opcoes["Opcao"].tolist())
    id_selecionado = str(opcao).split(" - ")[0].strip()

    idxs = tarefas.index[tarefas["ID"].astype(str).str.strip() == id_selecionado].tolist()
    if not idxs:
        st.error("Não consegui localizar essa tarefa no Excel.")
        return

    idx = idxs[0]
    atual = tarefas.loc[idx]

    deptos = get_departamentos(data) or ["Financeiro", "Controladoria", "Contabilidade", "Projetos"]
    if txt(atual.get("Departamento")) and txt(atual.get("Departamento")) not in deptos:
        deptos.append(txt(atual.get("Departamento")))

    usuarios = data["Usuarios"].copy()
    responsaveis = [""] + sorted([x for x in usuarios["Nome"].dropna().astype(str).unique() if x])
    if txt(atual.get("Responsavel")) and txt(atual.get("Responsavel")) not in responsaveis:
        responsaveis.append(txt(atual.get("Responsavel")))

    projetos = data["Projetos"].copy()
    projetos_l = [""] + sorted([x for x in projetos["Projeto"].dropna().astype(str).unique() if x])
    if txt(atual.get("Projeto")) and txt(atual.get("Projeto")) not in projetos_l:
        projetos_l.append(txt(atual.get("Projeto")))

    tarefas_l = [""] + sorted([x for x in tarefas["Tarefa"].dropna().astype(str).unique() if x and x != txt(atual.get("Tarefa"))])
    if txt(atual.get("Dependencia")) and txt(atual.get("Dependencia")) not in tarefas_l:
        tarefas_l.append(txt(atual.get("Dependencia")))

    periodicidades = ["Diario", "Semanal", "Mensal", "Unica"]
    if txt(atual.get("Periodicidade")) and txt(atual.get("Periodicidade")) not in periodicidades:
        periodicidades.append(txt(atual.get("Periodicidade")))

    prioridades = ["Normal", "Alta", "Crítica", "Baixa"]
    if txt(atual.get("Prioridade")) and txt(atual.get("Prioridade")) not in prioridades:
        prioridades.append(txt(atual.get("Prioridade")))

    status_opts = ["Pendente", "Concluída", "Suspensa"]
    if txt(atual.get("Status")) and txt(atual.get("Status")) not in status_opts:
        status_opts.append(txt(atual.get("Status")))

    data_ini_atual = parse_date(atual.get("Data de Inicio")) or date.today()

    with st.form("form_editar_tarefa"):
        c1, c2 = st.columns(2)

        with c1:
            tarefa = st.text_input("Tarefa", value=txt(atual.get("Tarefa")))
            desc = st.text_area("Descrição", value=txt(atual.get("Descrição")))
            depto = st.selectbox(
                "Departamento",
                deptos,
                index=deptos.index(txt(atual.get("Departamento"))) if txt(atual.get("Departamento")) in deptos else 0
            )
            projeto = st.selectbox(
                "Projeto",
                projetos_l,
                index=projetos_l.index(txt(atual.get("Projeto"))) if txt(atual.get("Projeto")) in projetos_l else 0
            )
            resp = st.selectbox(
                "Responsável",
                responsaveis,
                index=responsaveis.index(txt(atual.get("Responsavel"))) if txt(atual.get("Responsavel")) in responsaveis else 0
            )

        with c2:
            per = st.selectbox(
                "Periodicidade",
                periodicidades,
                index=periodicidades.index(txt(atual.get("Periodicidade"))) if txt(atual.get("Periodicidade")) in periodicidades else 0
            )
            obrig = st.selectbox(
                "Obrigatória",
                ["Não", "Sim"],
                index=1 if txt(atual.get("Obrigatoria")).lower() in ["sim", "s", "true", "1"] else 0
            )
            prio = st.selectbox(
                "Prioridade",
                prioridades,
                index=prioridades.index(txt(atual.get("Prioridade"))) if txt(atual.get("Prioridade")) in prioridades else 0
            )
            dep = st.selectbox(
                "Dependência",
                tarefas_l,
                index=tarefas_l.index(txt(atual.get("Dependencia"))) if txt(atual.get("Dependencia")) in tarefas_l else 0
            )
            data_ini = st.date_input("Data de início", value=data_ini_atual)
            status = st.selectbox(
                "Status",
                status_opts,
                index=status_opts.index(txt(atual.get("Status"))) if txt(atual.get("Status")) in status_opts else 0
            )

        obs = st.text_area("Observação", value=txt(atual.get("Observação")))

        cbtn1, cbtn2 = st.columns(2)
        with cbtn1:
            salvar = st.form_submit_button("💾 Salvar alterações")
        with cbtn2:
            arquivar = st.form_submit_button("🗄️ Arquivar tarefa")

    if salvar:
        tarefas = tarefas.astype("object")
        tarefas.at[idx, "Tarefa"] = str(tarefa)
        tarefas.at[idx, "Descrição"] = str(desc)
        tarefas.at[idx, "Departamento"] = normalizar_departamento(depto)
        tarefas.at[idx, "Projeto"] = str(projeto)
        tarefas.at[idx, "Responsavel"] = str(resp)
        tarefas.at[idx, "Periodicidade"] = str(per)
        tarefas.at[idx, "Obrigatoria"] = str(obrig)
        tarefas.at[idx, "Prioridade"] = str(prio)
        tarefas.at[idx, "Dependencia"] = str(dep)
        tarefas.at[idx, "Data de Inicio"] = data_ini.strftime("%d/%m/%Y")
        tarefas.at[idx, "Status"] = str(status)
        tarefas.at[idx, "Observação"] = str(obs)

        if status != "Concluída":
            tarefas.at[idx, "Concluído Por"] = ""
            tarefas.at[idx, "Data Conclusão"] = ""

        data["Tarefas"] = tarefas
        registrar_historico(data, id_selecionado, tarefa, user, "Alterada", "Tarefa editada pelo app")
        save_data(get_path(), data)
        st.success("Tarefa atualizada com sucesso.")
        st.rerun()

    if arquivar:
        tarefas = tarefas.astype("object")
        tarefas.at[idx, "Ativa"] = "Não"
        tarefas.at[idx, "Status"] = "Arquivada"
        data["Tarefas"] = tarefas
        registrar_historico(data, id_selecionado, txt(atual.get("Tarefa")), user, "Arquivada", "Tarefa arquivada pelo app")
        save_data(get_path(), data)
        st.success("Tarefa arquivada.")
        st.rerun()



def calendario_page(data, user):
    header("Calendário", "Visão operacional por data")
    tarefas = ensure_cols(data["Tarefas"], ABAS["Tarefas"])
    tarefas = tarefas[tarefas["Ativa"].apply(is_active)].reset_index(drop=True)
    tarefas["Departamento"] = tarefas["Departamento"].apply(normalizar_departamento)

    ref = st.date_input("Data", value=date.today())

    tarefas["ClassificacaoData"] = tarefas.apply(lambda r: classify(r, ref), axis=1)

    previstas_data = tarefas[
        (tarefas.apply(lambda r: periodic_on_date(r, ref), axis=1))
        & (~tarefas.apply(lambda r: done(r, ref), axis=1))
    ]

    concluidas_data = tarefas[tarefas.apply(lambda r: done(r, ref), axis=1)]

    c1, c2 = st.columns(2)
    with c1:
        metric_card("Previstas na data", len(previstas_data), ref.strftime("%d/%m/%Y"), "#f59e0b")
    with c2:
        metric_card("Concluídas na data", len(concluidas_data), "Somente nesta data", "#16a34a")

    st.markdown("<div class='panel'><div class='panel-title'>🟡 Tarefas previstas na data</div>", unsafe_allow_html=True)
    if previstas_data.empty:
        st.info("Nenhuma tarefa prevista/pendente para esta data.")
    else:
        st.dataframe(
            previstas_data[["ID", "Tarefa", "Departamento", "Responsavel", "Periodicidade", "Prioridade", "Status", "Data de Inicio"]],
            use_container_width=True,
            hide_index=True
        )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='panel'><div class='panel-title'>✅ Concluídas na data</div>", unsafe_allow_html=True)
    if concluidas_data.empty:
        st.info("Nenhuma tarefa concluída nesta data.")
    else:
        st.dataframe(
            concluidas_data[["ID", "Tarefa", "Departamento", "Responsavel", "Periodicidade", "Prioridade", "Status", "Data Conclusão"]],
            use_container_width=True,
            hide_index=True
        )
    st.markdown("</div>", unsafe_allow_html=True)


def historico_page(data):
    header("Histórico", "Registro das conclusões")
    hist = ensure_cols(data["Historico"], ABAS["Historico"])
    if hist.empty:
        st.info("Ainda não há histórico.")
    else:
        hist["ID Histórico Ordenacao"] = pd.to_numeric(hist["ID Histórico"], errors="coerce")
        hist = hist.sort_values("ID Histórico Ordenacao", ascending=False).drop(columns=["ID Histórico Ordenacao"])
        st.dataframe(hist, use_container_width=True, hide_index=True)


def main():
    data = load_data(get_path())
    user, page, departamentos = sidebar(data)

    if page == "Dashboard":
        dashboard(data, user)
    elif page == "Coordenação":
        coordenacao_page(data, user)
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

    voltar_topo()


if __name__ == "__main__":
    main()
