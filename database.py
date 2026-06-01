
# -*- coding: utf-8 -*-
import sqlite3
from pathlib import Path
from datetime import datetime, date
import pandas as pd


DB_PATH = "agenda.db"


def conectar(db_path=DB_PATH):
    conn = sqlite3.connect(db_path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def agora_br():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def hoje_br():
    return datetime.now().strftime("%d/%m/%Y")


def init_db(db_path=DB_PATH):
    conn = conectar(db_path)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE,
        login TEXT,
        senha TEXT,
        perfil TEXT,
        departamento TEXT,
        ativo TEXT DEFAULT 'Sim',
        data_cadastro TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS projetos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        projeto TEXT,
        objetivo TEXT,
        departamento TEXT,
        responsavel TEXT,
        data_inicio TEXT,
        prazo_final TEXT,
        status TEXT,
        percentual REAL DEFAULT 0,
        proxima_etapa TEXT,
        observacao TEXT,
        ativo TEXT DEFAULT 'Sim'
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS tarefas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tarefa TEXT NOT NULL,
        descricao TEXT,
        departamento TEXT,
        projeto TEXT,
        responsavel TEXT,
        periodicidade TEXT,
        obrigatoria TEXT DEFAULT 'Não',
        prioridade TEXT DEFAULT 'Normal',
        dependencia TEXT,
        prazo_limite TEXT,
        data_inicio TEXT,
        status TEXT DEFAULT 'Pendente',
        concluido_por TEXT,
        data_conclusao TEXT,
        observacao TEXT,
        ativa TEXT DEFAULT 'Sim',
        criado_por TEXT,
        data_criacao TEXT,
        ultima_atualizacao TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS historico (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        id_tarefa INTEGER,
        tarefa TEXT,
        usuario TEXT,
        data TEXT,
        hora TEXT,
        status TEXT,
        observacao TEXT,
        criado_em TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS configuracoes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT,
        valor TEXT,
        ativo TEXT DEFAULT 'Sim'
    )
    """)

    conn.commit()
    conn.close()


def executar(query, params=(), db_path=DB_PATH):
    conn = conectar(db_path)
    cur = conn.cursor()
    cur.execute(query, params)
    conn.commit()
    last_id = cur.lastrowid
    conn.close()
    return last_id


def consultar_df(query, params=(), db_path=DB_PATH):
    conn = conectar(db_path)
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def registrar_historico(id_tarefa, tarefa, usuario, status, observacao="", db_path=DB_PATH):
    now = datetime.now()
    executar(
        """
        INSERT INTO historico (id_tarefa, tarefa, usuario, data, hora, status, observacao, criado_em)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            id_tarefa,
            tarefa,
            usuario,
            now.strftime("%d/%m/%Y"),
            now.strftime("%H:%M:%S"),
            status,
            observacao,
            now.strftime("%d/%m/%Y %H:%M:%S"),
        ),
        db_path
    )


def listar_usuarios(db_path=DB_PATH):
    return consultar_df("SELECT * FROM usuarios WHERE ativo <> 'Não' ORDER BY nome", db_path=db_path)


def listar_departamentos(db_path=DB_PATH):
    df = consultar_df("""
        SELECT DISTINCT departamento 
        FROM tarefas 
        WHERE departamento IS NOT NULL AND TRIM(departamento) <> ''
        ORDER BY departamento
    """, db_path=db_path)
    return df["departamento"].dropna().tolist() if not df.empty else []


def listar_tarefas(ativas=True, db_path=DB_PATH):
    where = "WHERE ativa <> 'Não'" if ativas else ""
    return consultar_df(f"SELECT * FROM tarefas {where} ORDER BY id DESC", db_path=db_path)


def listar_historico(db_path=DB_PATH):
    return consultar_df("SELECT * FROM historico ORDER BY id DESC", db_path=db_path)


def listar_projetos(db_path=DB_PATH):
    return consultar_df("SELECT * FROM projetos WHERE ativo <> 'Não' ORDER BY id DESC", db_path=db_path)


def criar_usuario(nome, perfil="Usuário", departamento="", db_path=DB_PATH):
    if not nome:
        return
    executar(
        """
        INSERT OR IGNORE INTO usuarios (nome, perfil, departamento, ativo, data_cadastro)
        VALUES (?, ?, ?, 'Sim', ?)
        """,
        (nome, perfil, departamento, agora_br()),
        db_path
    )


def criar_tarefa(dados, usuario, db_path=DB_PATH):
    now = agora_br()
    id_tarefa = executar(
        """
        INSERT INTO tarefas (
            tarefa, descricao, departamento, projeto, responsavel, periodicidade,
            obrigatoria, prioridade, dependencia, prazo_limite, data_inicio,
            status, concluido_por, data_conclusao, observacao, ativa,
            criado_por, data_criacao, ultima_atualizacao
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Pendente', '', '', ?, 'Sim', ?, ?, ?)
        """,
        (
            dados.get("tarefa", ""),
            dados.get("descricao", ""),
            dados.get("departamento", ""),
            dados.get("projeto", ""),
            dados.get("responsavel", ""),
            dados.get("periodicidade", ""),
            dados.get("obrigatoria", "Não"),
            dados.get("prioridade", "Normal"),
            dados.get("dependencia", ""),
            dados.get("prazo_limite", ""),
            dados.get("data_inicio", ""),
            dados.get("observacao", ""),
            usuario,
            now,
            now,
        ),
        db_path
    )
    registrar_historico(id_tarefa, dados.get("tarefa", ""), usuario, "Criada", "Tarefa criada pelo app", db_path)
    return id_tarefa


def atualizar_status(id_tarefa, novo_status, usuario, observacao="", db_path=DB_PATH):
    tarefa_df = consultar_df("SELECT * FROM tarefas WHERE id = ?", (id_tarefa,), db_path)
    if tarefa_df.empty:
        return False

    tarefa = tarefa_df.iloc[0]["tarefa"]
    now = agora_br()

    concluido_por = usuario if novo_status == "Concluída" else ""
    data_conclusao = now if novo_status == "Concluída" else ""

    executar(
        """
        UPDATE tarefas
        SET status = ?, concluido_por = ?, data_conclusao = ?, ultima_atualizacao = ?
        WHERE id = ?
        """,
        (novo_status, concluido_por, data_conclusao, now, id_tarefa),
        db_path
    )

    registrar_historico(id_tarefa, tarefa, usuario, novo_status, observacao, db_path)
    return True


def adicionar_observacao(id_tarefa, usuario, observacao, db_path=DB_PATH):
    if not observacao:
        return False

    tarefa_df = consultar_df("SELECT * FROM tarefas WHERE id = ?", (id_tarefa,), db_path)
    if tarefa_df.empty:
        return False

    tarefa = tarefa_df.iloc[0]["tarefa"]
    now = agora_br()

    executar(
        "UPDATE tarefas SET observacao = ?, ultima_atualizacao = ? WHERE id = ?",
        (observacao, now, id_tarefa),
        db_path
    )

    registrar_historico(id_tarefa, tarefa, usuario, "Observação", observacao, db_path)
    return True


def editar_tarefa(id_tarefa, dados, usuario, db_path=DB_PATH):
    tarefa_df = consultar_df("SELECT * FROM tarefas WHERE id = ?", (id_tarefa,), db_path)
    if tarefa_df.empty:
        return False

    now = agora_br()

    executar(
        """
        UPDATE tarefas
        SET tarefa = ?, descricao = ?, departamento = ?, projeto = ?, responsavel = ?,
            periodicidade = ?, obrigatoria = ?, prioridade = ?, dependencia = ?,
            prazo_limite = ?, data_inicio = ?, status = ?, observacao = ?,
            ultima_atualizacao = ?
        WHERE id = ?
        """,
        (
            dados.get("tarefa", ""),
            dados.get("descricao", ""),
            dados.get("departamento", ""),
            dados.get("projeto", ""),
            dados.get("responsavel", ""),
            dados.get("periodicidade", ""),
            dados.get("obrigatoria", "Não"),
            dados.get("prioridade", "Normal"),
            dados.get("dependencia", ""),
            dados.get("prazo_limite", ""),
            dados.get("data_inicio", ""),
            dados.get("status", "Pendente"),
            dados.get("observacao", ""),
            now,
            id_tarefa,
        ),
        db_path
    )

    registrar_historico(id_tarefa, dados.get("tarefa", ""), usuario, "Alterada", "Tarefa editada pelo app", db_path)
    return True


def arquivar_tarefa(id_tarefa, usuario, db_path=DB_PATH):
    tarefa_df = consultar_df("SELECT * FROM tarefas WHERE id = ?", (id_tarefa,), db_path)
    if tarefa_df.empty:
        return False

    tarefa = tarefa_df.iloc[0]["tarefa"]

    executar(
        "UPDATE tarefas SET ativa = 'Não', status = 'Arquivada', ultima_atualizacao = ? WHERE id = ?",
        (agora_br(), id_tarefa),
        db_path
    )

    registrar_historico(id_tarefa, tarefa, usuario, "Arquivada", "Tarefa arquivada pelo app", db_path)
    return True


def reativar_tarefa(id_tarefa, usuario, db_path=DB_PATH):
    tarefa_df = consultar_df("SELECT * FROM tarefas WHERE id = ?", (id_tarefa,), db_path)
    if tarefa_df.empty:
        return False

    tarefa = tarefa_df.iloc[0]["tarefa"]

    executar(
        "UPDATE tarefas SET ativa = 'Sim', status = 'Pendente', ultima_atualizacao = ? WHERE id = ?",
        (agora_br(), id_tarefa),
        db_path
    )

    registrar_historico(id_tarefa, tarefa, usuario, "Reativada", "Tarefa reativada pelo app", db_path)
    return True


def exportar_excel(caminho_saida, db_path=DB_PATH):
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        consultar_df("SELECT * FROM usuarios", db_path=db_path).to_excel(writer, sheet_name="Usuarios", index=False)
        consultar_df("SELECT * FROM tarefas", db_path=db_path).to_excel(writer, sheet_name="Tarefas", index=False)
        consultar_df("SELECT * FROM projetos", db_path=db_path).to_excel(writer, sheet_name="Projetos", index=False)
        consultar_df("SELECT * FROM historico", db_path=db_path).to_excel(writer, sheet_name="Historico", index=False)
        consultar_df("SELECT * FROM configuracoes", db_path=db_path).to_excel(writer, sheet_name="Configuracoes", index=False)
