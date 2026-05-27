
# -*- coding: utf-8 -*-
import os
from pathlib import Path
from datetime import datetime, date, time, timedelta

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================

st.set_page_config(
    page_title="Agenda Departamental",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARQUIVO_EXCEL_PADRAO = "Agenda.xlsx"

ABAS_OBRIGATORIAS = {
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
# ESTILO VISUAL
# ============================================================

st.markdown(
    """
    <style>
        .main {
            background-color: #f7f9fb;
        }
        div[data-testid="stMetric"] {
            background: white;
            border: 1px solid #e5e7eb;
            padding: 16px;
            border-radius: 18px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        .task-card {
            background: white;
            padding: 16px;
            border-radius: 18px;
            border: 1px solid #e5e7eb;
            margin-bottom: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        }
        .status-vencida {
            color: #b91c1c;
            font-weight: 700;
        }
        .status-hoje {
            color: #92400e;
            font-weight: 700;
        }
        .status-ok {
            color: #166534;
            font-weight: 700;
        }
        .small-muted {
            color: #6b7280;
            font-size: 13px;
        }
        .stButton > button {
            border-radius: 12px;
            font-weight: 600;
            width: 100%;
        }
        .title-box {
            padding: 18px 20px;
            background: linear-gradient(90deg, #0f172a, #1e3a8a);
            color: white;
            border-radius: 20px;
            margin-bottom: 18px;
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def sim(valor):
    return normalizar_texto(valor).lower() in ["sim", "s", "yes", "y", "true", "1", "ativo", "ativa"]


def hoje_data():
    return date.today()


def agora():
    return datetime.now()


def parse_data(valor):
    if pd.isna(valor) or valor == "":
        return None
    try:
        return pd.to_datetime(valor, dayfirst=True).date()
    except Exception:
        return None


def parse_hora_limite(valor):
    if pd.isna(valor) or valor == "":
        return None

    if isinstance(valor, time):
        return valor

    texto = str(valor).strip()
    if not texto:
        return None

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(texto, fmt).time()
        except Exception:
            pass

    return None


def proximo_id(df, coluna):
    if df.empty or coluna not in df.columns:
        return 1
    ids = pd.to_numeric(df[coluna], errors="coerce").dropna()
    if ids.empty:
        return 1
    return int(ids.max()) + 1


def garantir_colunas(df, colunas):
    df = df.copy()
    for col in colunas:
        if col not in df.columns:
            df[col] = ""
    return df[colunas]


def status_ativo(valor):
    texto = normalizar_texto(valor).lower()
    if texto in ["não", "nao", "n", "no", "false", "0", "inativo", "inativa", "arquivada", "arquivado"]:
        return False
    return True


def tarefa_concluida_hoje(row):
    status = normalizar_texto(row.get("Status")).lower()
    data_conclusao = parse_data(row.get("Data Conclusão"))
    return status in ["concluída", "concluida", "finalizada", "feito"] and data_conclusao == hoje_data()


def tarefa_concluida(row):
    status = normalizar_texto(row.get("Status")).lower()
    return status in ["concluída", "concluida", "finalizada", "feito"]


def prazo_vencido(row):
    if tarefa_concluida(row):
        return False

    data_inicio = parse_data(row.get("Data de Inicio"))
    periodicidade = normalizar_texto(row.get("Periodicidade")).lower()

    if data_inicio and data_inicio > hoje_data():
        return False

    # Se não houver data de início, consideramos como rotina ativa.
    if not data_inicio:
        data_inicio = hoje_data()

    if periodicidade in ["diario", "diária", "diaria", "todo dia", "diário"]:
        return False

    # Para tarefas pontuais ou com data inicial antiga, considera vencida.
    if data_inicio and data_inicio < hoje_data():
        return True

    return False


def tarefa_para_hoje(row):
    if tarefa_concluida(row):
        return False

    if not status_ativo(row.get("Ativa", "Sim")):
        return False

    data_inicio = parse_data(row.get("Data de Inicio"))
    periodicidade = normalizar_texto(row.get("Periodicidade")).lower()

    if data_inicio and data_inicio > hoje_data():
        return False

    if periodicidade in ["", "diario", "diária", "diaria", "diário", "todo dia"]:
        return True

    if periodicidade in ["semanal", "semana"]:
        if data_inicio:
            return (hoje_data() - data_inicio).days % 7 == 0
        return True

    if periodicidade in ["mensal", "mês", "mes"]:
        if data_inicio:
            return hoje_data().day == data_inicio.day
        return True

    if periodicidade in ["unica", "única", "pontual"]:
        return data_inicio == hoje_data()

    return True


def classificar_tarefa(row):
    if tarefa_concluida(row):
        return "Concluída"
    if prazo_vencido(row):
        return "Vencida"
    if tarefa_para_hoje(row):
        return "Hoje"
    return "Futura"


def dependencias_pendentes(row, tarefas):
    dependencia = normalizar_texto(row.get("Dependencia"))
    if not dependencia:
        return ""

    nomes = [x.strip() for x in dependencia.replace(";", ",").split(",") if x.strip()]
    if not nomes:
        return ""

    pendentes = []
    for nome in nomes:
        filtro = tarefas["Tarefa"].astype(str).str.strip().str.lower() == nome.lower()
        encontrados = tarefas[filtro]
        if encontrados.empty:
            pendentes.append(f"{nome} não localizada")
        else:
            if not encontrados.apply(tarefa_concluida, axis=1).any():
                pendentes.append(nome)

    return ", ".join(pendentes)


# ============================================================
# LEITURA E SALVAMENTO
# ============================================================

@st.cache_data(show_spinner=False)
def carregar_excel(caminho):
    caminho = Path(caminho)

    if not caminho.exists():
        dados = {aba: pd.DataFrame(columns=cols) for aba, cols in ABAS_OBRIGATORIAS.items()}
        return dados

    dados = {}
    xls = pd.ExcelFile(caminho)
    for aba, colunas in ABAS_OBRIGATORIAS.items():
        if aba in xls.sheet_names:
            df = pd.read_excel(caminho, sheet_name=aba)
        else:
            df = pd.DataFrame(columns=colunas)
        dados[aba] = garantir_colunas(df, colunas)

    return dados


def salvar_excel(caminho, dados):
    caminho = Path(caminho)
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for aba, colunas in ABAS_OBRIGATORIAS.items():
            df = dados.get(aba, pd.DataFrame(columns=colunas))
            df = garantir_colunas(df, colunas)
            df.to_excel(writer, sheet_name=aba, index=False)

    st.cache_data.clear()


def obter_caminho_excel():
    if "arquivo_excel" not in st.session_state:
        st.session_state["arquivo_excel"] = ARQUIVO_EXCEL_PADRAO
    return st.session_state["arquivo_excel"]


def carregar_dados_atuais():
    return carregar_excel(obter_caminho_excel())


# ============================================================
# COMPONENTES DE INTERFACE
# ============================================================

def cabecalho():
    st.markdown(
        """
        <div class="title-box">
            <h2 style="margin:0;">📅 Agenda Departamental</h2>
            <p style="margin:6px 0 0 0;">Rotinas, pendências, projetos e calendário operacional</p>
        </div>
        """,
        unsafe_allow_html=True
    )


def sidebar_config(dados):
    st.sidebar.title("⚙️ Configuração")

    caminho_atual = obter_caminho_excel()
    caminho = st.sidebar.text_input("Arquivo Excel", value=caminho_atual)

    if caminho != caminho_atual:
        st.session_state["arquivo_excel"] = caminho
        st.cache_data.clear()
        st.rerun()

    usuarios = dados["Usuarios"].copy()
    usuarios_ativos = usuarios[usuarios["Ativo"].apply(status_ativo)] if not usuarios.empty else usuarios

    nomes = sorted([normalizar_texto(x) for x in usuarios_ativos["Nome"].dropna().unique() if normalizar_texto(x)])
    if not nomes:
        nomes = ["Paula"]

    usuario = st.sidebar.selectbox("Quem está usando?", nomes)
    st.session_state["usuario_atual"] = usuario

    st.sidebar.caption("Nesta V1 todos visualizam tudo. O usuário selecionado é usado para registrar histórico.")

    if st.sidebar.button("🔄 Atualizar dados"):
        st.cache_data.clear()
        st.rerun()

    return usuario


def filtro_base(tarefas):
    tarefas = tarefas.copy()
    tarefas = tarefas[tarefas["Ativa"].apply(status_ativo)]
    tarefas["Classificação"] = tarefas.apply(classificar_tarefa, axis=1)
    return tarefas


def card_tarefa(row, dados, usuario, key_prefix="task"):
    tarefas = dados["Tarefas"].copy()
    historico = dados["Historico"].copy()

    id_tarefa = row.get("ID")
    tarefa = normalizar_texto(row.get("Tarefa"))
    desc = normalizar_texto(row.get("Descrição"))
    departamento = normalizar_texto(row.get("Departamento"))
    responsavel = normalizar_texto(row.get("Responsavel"))
    prioridade = normalizar_texto(row.get("Prioridade")) or "Normal"
    projeto = normalizar_texto(row.get("Projeto"))
    classificacao = normalizar_texto(row.get("Classificação"))
    pendencias_dep = dependencias_pendentes(row, tarefas)

    status_class = "status-ok" if classificacao == "Concluída" else "status-vencida" if classificacao == "Vencida" else "status-hoje"
    concluida = tarefa_concluida(row)

    with st.container(border=True):
        col_info, col_meta, col_acao = st.columns([5, 2, 2])

        with col_info:
            st.markdown(f"**{tarefa}**")
            if desc:
                st.caption(desc)
            if projeto:
                st.caption(f"Projeto: {projeto}")
            if pendencias_dep:
                st.warning(f"Dependência pendente: {pendencias_dep}")

        with col_meta:
            st.markdown(f"<span class='{status_class}'>{classificacao}</span>", unsafe_allow_html=True)
            st.caption(f"Departamento: {departamento or '-'}")
            st.caption(f"Responsável: {responsavel or '-'}")
            st.caption(f"Prioridade: {prioridade}")

        with col_acao:
            if concluida:
                st.success("✅ Concluída")
                if st.button("↩️ Reabrir", key=f"reabrir_{key_prefix}_{id_tarefa}"):
                    idx_list = tarefas.index[tarefas["ID"].astype(str) == str(id_tarefa)].tolist()
                    if idx_list:
                        idx = idx_list[0]
                        tarefas.at[idx, "Status"] = "Pendente"
                        tarefas.at[idx, "Concluído Por"] = ""
                        tarefas.at[idx, "Data Conclusão"] = ""
                        dados["Tarefas"] = tarefas
                        salvar_excel(obter_caminho_excel(), dados)
                        st.rerun()
            else:
                if st.button("✅ Marcar concluída", key=f"concluir_{key_prefix}_{id_tarefa}"):
                    idx_list = tarefas.index[tarefas["ID"].astype(str) == str(id_tarefa)].tolist()

                    if not idx_list:
                        st.error("Não consegui localizar esta tarefa no Excel.")
                        return

                    idx = idx_list[0]
                    agora_dt = agora()

                    tarefas.at[idx, "Status"] = "Concluída"
                    tarefas.at[idx, "Concluído Por"] = usuario
                    tarefas.at[idx, "Data Conclusão"] = agora_dt.strftime("%d/%m/%Y %H:%M:%S")

                    novo_hist = {
                        "ID Histórico": proximo_id(historico, "ID Histórico"),
                        "ID Tarefa": id_tarefa,
                        "Tarefa": tarefa,
                        "Usuário": usuario,
                        "Data": agora_dt.strftime("%d/%m/%Y"),
                        "Hora": agora_dt.strftime("%H:%M:%S"),
                        "Status": "Concluída",
                        "Observação": "",
                    }

                    historico = pd.concat([historico, pd.DataFrame([novo_hist])], ignore_index=True)

                    dados["Tarefas"] = tarefas
                    dados["Historico"] = historico
                    salvar_excel(obter_caminho_excel(), dados)

                    st.success("Tarefa concluída e histórico registrado.")
                    st.rerun()



# ============================================================
# PÁGINAS
# ============================================================

def pagina_dashboard(dados, usuario):
    cabecalho()

    tarefas = filtro_base(dados["Tarefas"])
    projetos = dados["Projetos"].copy()
    projetos_ativos = projetos[projetos["Ativo"].apply(status_ativo)] if not projetos.empty else projetos

    total_hoje = (tarefas["Classificação"] == "Hoje").sum()
    total_vencidas = (tarefas["Classificação"] == "Vencida").sum()
    total_concluidas = (tarefas["Classificação"] == "Concluída").sum()
    total_projetos = len(projetos_ativos)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🟡 Tarefas de hoje", int(total_hoje))
    c2.metric("🔴 Pendências", int(total_vencidas))
    c3.metric("🟢 Concluídas", int(total_concluidas))
    c4.metric("📌 Projetos ativos", int(total_projetos))

    st.divider()

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("🔴 Pendências anteriores")
        pendentes = tarefas[tarefas["Classificação"] == "Vencida"]
        if pendentes.empty:
            st.success("Nenhuma pendência anterior.")
        else:
            for _, row in pendentes.iterrows():
                card_tarefa(row, dados, usuario, key_prefix="dash_vencida")

    with col2:
        st.subheader("🟡 Agenda de hoje")
        hoje = tarefas[tarefas["Classificação"] == "Hoje"]
        if hoje.empty:
            st.info("Nenhuma tarefa prevista para hoje.")
        else:
            for _, row in hoje.iterrows():
                card_tarefa(row, dados, usuario, key_prefix="dash_hoje")


def pagina_departamento(dados, usuario, departamento):
    cabecalho()
    st.subheader(f"📂 {departamento}")

    tarefas = filtro_base(dados["Tarefas"])
    tarefas = tarefas[tarefas["Departamento"].astype(str).str.strip().str.lower() == departamento.lower()]

    if tarefas.empty:
        st.info("Nenhuma tarefa cadastrada para este departamento.")
        return

    status_filtro = st.multiselect(
        "Filtrar por status",
        options=["Vencida", "Hoje", "Futura", "Concluída"],
        default=["Vencida", "Hoje", "Futura"],
        key=f"filtro_{departamento}"
    )

    busca = st.text_input("Buscar tarefa", key=f"busca_{departamento}")

    filtradas = tarefas[tarefas["Classificação"].isin(status_filtro)]
    if busca:
        busca_l = busca.lower()
        filtradas = filtradas[
            filtradas["Tarefa"].astype(str).str.lower().str.contains(busca_l, na=False)
            | filtradas["Descrição"].astype(str).str.lower().str.contains(busca_l, na=False)
            | filtradas["Responsavel"].astype(str).str.lower().str.contains(busca_l, na=False)
        ]

    for _, row in filtradas.iterrows():
        card_tarefa(row, dados, usuario, key_prefix=f"dep_{departamento}")


def pagina_projetos(dados, usuario):
    cabecalho()
    st.subheader("📌 Projetos")

    projetos = dados["Projetos"].copy()
    projetos = projetos[projetos["Ativo"].apply(status_ativo)] if not projetos.empty else projetos

    if projetos.empty:
        st.info("Nenhum projeto cadastrado.")
    else:
        for _, row in projetos.iterrows():
            projeto = normalizar_texto(row.get("Projeto"))
            objetivo = normalizar_texto(row.get("Objetivo"))
            responsavel = normalizar_texto(row.get("Responsavel"))
            departamento = normalizar_texto(row.get("Departamento"))
            status = normalizar_texto(row.get("Status")) or "Em andamento"
            prazo = normalizar_texto(row.get("Prazo Final"))
            try:
                percentual = float(row.get("%") or 0)
            except Exception:
                percentual = 0

            with st.container(border=True):
                c1, c2 = st.columns([4, 1])
                with c1:
                    st.markdown(f"### {projeto}")
                    if objetivo:
                        st.write(objetivo)
                    st.caption(f"Departamento: {departamento or '-'} | Responsável: {responsavel or '-'}")
                    st.caption(f"Prazo: {prazo or '-'} | Status: {status}")
                    st.progress(min(max(percentual / 100, 0), 1))
                with c2:
                    st.metric("% concluído", f"{percentual:.0f}%")

    st.divider()
    st.subheader("➕ Cadastrar novo projeto")

    with st.form("form_novo_projeto"):
        projeto = st.text_input("Projeto")
        objetivo = st.text_area("Objetivo")
        departamento = st.text_input("Departamento")
        responsavel = st.text_input("Responsável")
        prazo = st.date_input("Prazo final", value=None)
        status = st.selectbox("Status", ["Em andamento", "Pendente", "Concluído", "Suspenso"])
        percentual = st.slider("% concluído", 0, 100, 0)
        proxima_etapa = st.text_input("Próxima etapa")
        observacao = st.text_area("Observação")
        salvar = st.form_submit_button("Salvar projeto")

    if salvar:
        if not projeto:
            st.error("Informe o nome do projeto.")
        else:
            novo = {
                "ID Projeto": proximo_id(dados["Projetos"], "ID Projeto"),
                "Projeto": projeto,
                "Objetivo": objetivo,
                "Departamento": departamento,
                "Responsavel": responsavel,
                "Data de Inicio": hoje_data().strftime("%d/%m/%Y"),
                "Prazo Final": prazo.strftime("%d/%m/%Y") if prazo else "",
                "Status": status,
                "%": percentual,
                "Proxima Etapa": proxima_etapa,
                "Observação": observacao,
                "Ativo": "Sim",
            }
            dados["Projetos"] = pd.concat([dados["Projetos"], pd.DataFrame([novo])], ignore_index=True)
            salvar_excel(obter_caminho_excel(), dados)
            st.success("Projeto cadastrado.")
            st.rerun()


def pagina_cadastro_tarefa(dados, usuario):
    cabecalho()
    st.subheader("➕ Cadastro de tarefas")

    tarefas = dados["Tarefas"]
    projetos = dados["Projetos"]
    usuarios = dados["Usuarios"]

    departamentos_existentes = sorted([
        x for x in tarefas["Departamento"].dropna().astype(str).str.strip().unique() if x
    ])
    if not departamentos_existentes:
        departamentos_existentes = ["Financeiro", "Controladoria", "Contabilidade", "Projetos"]

    projetos_lista = [""] + sorted([
        x for x in projetos["Projeto"].dropna().astype(str).str.strip().unique() if x
    ])

    responsaveis = [""] + sorted([
        x for x in usuarios["Nome"].dropna().astype(str).str.strip().unique() if x
    ])

    tarefas_lista = [""] + sorted([
        x for x in tarefas["Tarefa"].dropna().astype(str).str.strip().unique() if x
    ])

    with st.form("form_tarefa"):
        col1, col2 = st.columns(2)

        with col1:
            tarefa = st.text_input("Tarefa")
            descricao = st.text_area("Descrição / Subitem")
            departamento = st.selectbox("Departamento", departamentos_existentes)
            projeto = st.selectbox("Projeto vinculado", projetos_lista)
            responsavel = st.selectbox("Responsável", responsaveis)

        with col2:
            periodicidade = st.selectbox("Periodicidade", ["Diario", "Semanal", "Mensal", "Unica"])
            obrigatoria = st.selectbox("Obrigatória", ["Não", "Sim"])
            prioridade = st.selectbox("Prioridade", ["Normal", "Alta", "Crítica", "Baixa"])
            dependencia = st.selectbox("Dependência", tarefas_lista)
            data_inicio = st.date_input("Data de início", value=hoje_data())
            prazo_limite = st.time_input("Prazo limite", value=time(18, 0))

        observacao = st.text_area("Observação")
        salvar = st.form_submit_button("Salvar tarefa")

    if salvar:
        if not tarefa:
            st.error("Informe o nome da tarefa.")
        else:
            novo = {
                "ID": proximo_id(tarefas, "ID"),
                "Tarefa": tarefa,
                "Descrição": descricao,
                "Departamento": departamento,
                "Projeto": projeto,
                "Responsavel": responsavel,
                "Periodicidade": periodicidade,
                "Obrigatoria": obrigatoria,
                "Prioridade": prioridade,
                "Dependencia": dependencia,
                "Prazo Limite": prazo_limite.strftime("%H:%M:%S") if prazo_limite else "",
                "Data de Inicio": data_inicio.strftime("%d/%m/%Y") if data_inicio else "",
                "Status": "Pendente",
                "Concluído Por": "",
                "Data Conclusão": "",
                "Observação": observacao,
                "Ativa": "Sim",
            }
            dados["Tarefas"] = pd.concat([tarefas, pd.DataFrame([novo])], ignore_index=True)
            salvar_excel(obter_caminho_excel(), dados)
            st.success("Tarefa cadastrada.")
            st.rerun()

    st.divider()
    st.subheader("📋 Tarefas cadastradas")

    visual = dados["Tarefas"].copy()
    visual["Classificação"] = visual.apply(classificar_tarefa, axis=1)
    st.dataframe(visual, use_container_width=True, hide_index=True)


def pagina_historico(dados):
    cabecalho()
    st.subheader("🧾 Histórico")

    historico = dados["Historico"].copy()
    if historico.empty:
        st.info("Ainda não há registros no histórico.")
        return

    busca = st.text_input("Buscar no histórico")
    if busca:
        busca_l = busca.lower()
        historico = historico[
            historico["Tarefa"].astype(str).str.lower().str.contains(busca_l, na=False)
            | historico["Usuário"].astype(str).str.lower().str.contains(busca_l, na=False)
        ]

    st.dataframe(historico.sort_values(by="ID Histórico", ascending=False), use_container_width=True, hide_index=True)


def pagina_calendario(dados, usuario):
    cabecalho()
    st.subheader("🗓️ Calendário operacional")

    tarefas = filtro_base(dados["Tarefas"])

    col1, col2 = st.columns([1, 2])
    with col1:
        data_ref = st.date_input("Data de referência", value=hoje_data())

    with col2:
        st.info("Nesta V1, o calendário mostra tarefas ativas da data selecionada e pendências em aberto.")

    tarefas_data = tarefas.copy()
    tarefas_data["DataBase"] = tarefas_data["Data de Inicio"].apply(parse_data)

    vencidas = tarefas_data[
        (~tarefas_data.apply(tarefa_concluida, axis=1))
        & (tarefas_data["DataBase"].notna())
        & (tarefas_data["DataBase"] < data_ref)
    ]

    do_dia = tarefas_data[
        (~tarefas_data.apply(tarefa_concluida, axis=1))
        & (
            tarefas_data["DataBase"].isna()
            | (tarefas_data["DataBase"] <= data_ref)
        )
    ]

    st.markdown("### 🔴 Pendências até a data")
    if vencidas.empty:
        st.success("Sem pendências anteriores para esta data.")
    else:
        st.dataframe(vencidas[["ID", "Tarefa", "Departamento", "Responsavel", "Prioridade", "Data de Inicio", "Status"]], use_container_width=True, hide_index=True)

    st.markdown("### 🟡 Rotina/agenda da data")
    if do_dia.empty:
        st.info("Nenhuma tarefa encontrada para esta data.")
    else:
        st.dataframe(do_dia[["ID", "Tarefa", "Departamento", "Responsavel", "Periodicidade", "Prioridade", "Status"]], use_container_width=True, hide_index=True)


# ============================================================
# MAIN
# ============================================================

def main():
    dados = carregar_dados_atuais()
    usuario = sidebar_config(dados)

    st.sidebar.divider()
    pagina = st.sidebar.radio(
        "Menu",
        [
            "Dashboard",
            "Financeiro",
            "Controladoria",
            "Contabilidade",
            "Projetos",
            "Cadastro de tarefas",
            "Calendário",
            "Histórico",
        ],
    )

    if pagina == "Dashboard":
        pagina_dashboard(dados, usuario)
    elif pagina in ["Financeiro", "Controladoria", "Contabilidade"]:
        pagina_departamento(dados, usuario, pagina)
    elif pagina == "Projetos":
        pagina_projetos(dados, usuario)
    elif pagina == "Cadastro de tarefas":
        pagina_cadastro_tarefa(dados, usuario)
    elif pagina == "Calendário":
        pagina_calendario(dados, usuario)
    elif pagina == "Histórico":
        pagina_historico(dados)


if __name__ == "__main__":
    main()
