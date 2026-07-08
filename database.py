# -*- coding: utf-8 -*-
import sqlite3
from datetime import datetime
import pandas as pd

DB_PATH = "agenda_v20.db"

def conectar(db_path=DB_PATH):
    conn = sqlite3.connect(db_path, check_same_thread=False, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn

def agora():
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")

def executar(query, params=(), db_path=DB_PATH):
    conn = conectar(db_path); cur=conn.cursor(); cur.execute(query, params); conn.commit(); lid=cur.lastrowid; conn.close(); return lid

def consultar_df(query, params=(), db_path=DB_PATH):
    conn=conectar(db_path); df=pd.read_sql_query(query, conn, params=params); conn.close(); return df

def init_db(db_path=DB_PATH):
    conn=conectar(db_path); cur=conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY AUTOINCREMENT,nome TEXT UNIQUE,perfil TEXT DEFAULT 'Usuário',departamento TEXT,ativo TEXT DEFAULT 'Sim',criado_em TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS rotinas (id INTEGER PRIMARY KEY AUTOINCREMENT,rotina TEXT NOT NULL,descricao TEXT,departamento TEXT,responsavel TEXT,periodicidade TEXT DEFAULT 'Diaria',obrigatoria TEXT DEFAULT 'Não',prioridade TEXT DEFAULT 'Normal',dia_semana TEXT,dia_mes TEXT,data_inicio TEXT,projeto TEXT,ativa TEXT DEFAULT 'Sim',criado_por TEXT,criado_em TEXT,ultima_atualizacao TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS tarefas_dia (id INTEGER PRIMARY KEY AUTOINCREMENT,rotina_id INTEGER,data_ref TEXT,tarefa TEXT,descricao TEXT,departamento TEXT,responsavel TEXT,projeto TEXT,obrigatoria TEXT,prioridade TEXT,status TEXT DEFAULT 'Pendente',observacao TEXT,concluido_por TEXT,concluido_em TEXT,reprogramada_para TEXT,cancelada_motivo TEXT,criado_em TEXT,ultima_atualizacao TEXT,UNIQUE(rotina_id,data_ref))""")
    cur.execute("""CREATE TABLE IF NOT EXISTS projetos (id INTEGER PRIMARY KEY AUTOINCREMENT,projeto TEXT UNIQUE,objetivo TEXT,departamento TEXT,responsavel TEXT,data_inicio TEXT,prazo_final TEXT,status TEXT DEFAULT 'Em andamento',proxima_etapa TEXT,observacao TEXT,ativo TEXT DEFAULT 'Sim',criado_em TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS historico (id INTEGER PRIMARY KEY AUTOINCREMENT,entidade TEXT,entidade_id INTEGER,data_ref TEXT,usuario TEXT,acao TEXT,observacao TEXT,criado_em TEXT)""")
    conn.commit(); conn.close()

def registrar_historico(entidade, entidade_id, data_ref, usuario, acao, observacao=""):
    executar("INSERT INTO historico (entidade,entidade_id,data_ref,usuario,acao,observacao,criado_em) VALUES (?,?,?,?,?,?,?)", (entidade, entidade_id, data_ref, usuario, acao, observacao, agora()))

def listar_usuarios(): return consultar_df("SELECT * FROM usuarios WHERE ativo <> 'Não' ORDER BY nome")
def criar_usuario(nome, perfil='Usuário', departamento=''):
    if nome: executar("INSERT OR IGNORE INTO usuarios (nome,perfil,departamento,ativo,criado_em) VALUES (?,?,?,'Sim',?)", (nome, perfil, departamento, agora()))

def listar_departamentos():
    df=consultar_df("SELECT DISTINCT departamento FROM rotinas WHERE ativa <> 'Não' AND departamento IS NOT NULL AND TRIM(departamento)<>'' ORDER BY departamento")
    return df['departamento'].dropna().astype(str).tolist() if not df.empty else []

def listar_rotinas(ativas=True): return consultar_df(f"SELECT * FROM rotinas {'WHERE ativa <> \'Não\'' if ativas else ''} ORDER BY departamento, rotina")

def criar_rotina(dados, usuario):
    rid=executar("""INSERT INTO rotinas (rotina,descricao,departamento,responsavel,periodicidade,obrigatoria,prioridade,dia_semana,dia_mes,data_inicio,projeto,ativa,criado_por,criado_em,ultima_atualizacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,'Sim',?,?,?)""", (dados.get('rotina',''),dados.get('descricao',''),dados.get('departamento',''),dados.get('responsavel',''),dados.get('periodicidade','Diaria'),dados.get('obrigatoria','Não'),dados.get('prioridade','Normal'),dados.get('dia_semana',''),dados.get('dia_mes',''),dados.get('data_inicio',''),dados.get('projeto',''),usuario,agora(),agora()))
    registrar_historico('rotina', rid, '', usuario, 'Rotina criada', dados.get('rotina','')); return rid

def arquivar_rotina(rotina_id, usuario):
    executar("UPDATE rotinas SET ativa='Não', ultima_atualizacao=? WHERE id=?", (agora(), rotina_id)); registrar_historico('rotina', rotina_id, '', usuario, 'Rotina arquivada', '')

def listar_tarefas_dia(data_ref=None, incluir_todas=False):
    if data_ref and not incluir_todas: return consultar_df("SELECT * FROM tarefas_dia WHERE data_ref=? ORDER BY departamento,responsavel,tarefa", (data_ref,))
    return consultar_df("SELECT * FROM tarefas_dia ORDER BY data_ref DESC, departamento, tarefa")

def listar_pendencias_anteriores(data_ref):
    return consultar_df("SELECT * FROM tarefas_dia WHERE data_ref < ? AND status NOT IN ('Concluída','Cancelada','Reprogramada') ORDER BY data_ref,departamento,responsavel,tarefa", (data_ref,))

def criar_tarefa_dia(rotina, data_ref):
    executar("""INSERT OR IGNORE INTO tarefas_dia (rotina_id,data_ref,tarefa,descricao,departamento,responsavel,projeto,obrigatoria,prioridade,status,observacao,concluido_por,concluido_em,reprogramada_para,cancelada_motivo,criado_em,ultima_atualizacao) VALUES (?,?,?,?,?,?,?,?,?,'Pendente','','','','','',?,?)""", (int(rotina['id']),data_ref,rotina['rotina'],rotina.get('descricao',''),rotina.get('departamento',''),rotina.get('responsavel',''),rotina.get('projeto',''),rotina.get('obrigatoria','Não'),rotina.get('prioridade','Normal'),agora(),agora()))

def atualizar_status_tarefa(tarefa_id, status, usuario, observacao=''):
    concluido_por=usuario if status=='Concluída' else ''; concluido_em=agora() if status=='Concluída' else ''
    executar("UPDATE tarefas_dia SET status=?, observacao=?, concluido_por=?, concluido_em=?, ultima_atualizacao=? WHERE id=?", (status, observacao, concluido_por, concluido_em, agora(), tarefa_id))
    df=consultar_df("SELECT data_ref FROM tarefas_dia WHERE id=?", (tarefa_id,)); data_ref=df.iloc[0]['data_ref'] if not df.empty else ''
    registrar_historico('tarefa_dia', tarefa_id, data_ref, usuario, status, observacao)

def adicionar_observacao_tarefa(tarefa_id, usuario, observacao):
    executar("UPDATE tarefas_dia SET observacao=?, ultima_atualizacao=? WHERE id=?", (observacao, agora(), tarefa_id))
    df=consultar_df("SELECT data_ref FROM tarefas_dia WHERE id=?", (tarefa_id,)); data_ref=df.iloc[0]['data_ref'] if not df.empty else ''
    registrar_historico('tarefa_dia', tarefa_id, data_ref, usuario, 'Observação', observacao)

def reprogramar_tarefa(tarefa_id, usuario, nova_data, motivo):
    executar("UPDATE tarefas_dia SET status='Reprogramada', reprogramada_para=?, observacao=?, ultima_atualizacao=? WHERE id=?", (nova_data, motivo, agora(), tarefa_id))
    df=consultar_df("SELECT * FROM tarefas_dia WHERE id=?", (tarefa_id,))
    if not df.empty:
        r=df.iloc[0].to_dict(); registrar_historico('tarefa_dia', tarefa_id, r.get('data_ref',''), usuario, 'Reprogramada', motivo)
        executar("""INSERT OR IGNORE INTO tarefas_dia (rotina_id,data_ref,tarefa,descricao,departamento,responsavel,projeto,obrigatoria,prioridade,status,observacao,criado_em,ultima_atualizacao) VALUES (?,?,?,?,?,?,?,?,?,'Pendente',?,?,?)""", (r.get('rotina_id'), nova_data, r.get('tarefa'), r.get('descricao'), r.get('departamento'), r.get('responsavel'), r.get('projeto'), r.get('obrigatoria'), r.get('prioridade'), f"Reprogramada de {r.get('data_ref')}: {motivo}", agora(), agora()))

def cancelar_tarefa(tarefa_id, usuario, motivo):
    executar("UPDATE tarefas_dia SET status='Cancelada', cancelada_motivo=?, observacao=?, ultima_atualizacao=? WHERE id=?", (motivo, motivo, agora(), tarefa_id))
    df=consultar_df("SELECT data_ref FROM tarefas_dia WHERE id=?", (tarefa_id,)); data_ref=df.iloc[0]['data_ref'] if not df.empty else ''
    registrar_historico('tarefa_dia', tarefa_id, data_ref, usuario, 'Cancelada', motivo)

def listar_projetos(): return consultar_df("SELECT * FROM projetos WHERE ativo <> 'Não' ORDER BY projeto")
def criar_projeto(dados, usuario):
    pid=executar("""INSERT OR IGNORE INTO projetos (projeto,objetivo,departamento,responsavel,data_inicio,prazo_final,status,proxima_etapa,observacao,ativo,criado_em) VALUES (?,?,?,?,?,?,?,?,?,'Sim',?)""", (dados.get('projeto',''),dados.get('objetivo',''),dados.get('departamento',''),dados.get('responsavel',''),dados.get('data_inicio',''),dados.get('prazo_final',''),dados.get('status','Em andamento'),dados.get('proxima_etapa',''),dados.get('observacao',''),agora()))
    registrar_historico('projeto', pid, '', usuario, 'Projeto criado', dados.get('projeto','')); return pid

def listar_historico(): return consultar_df("SELECT * FROM historico ORDER BY id DESC")
def exportar_excel(caminho_saida):
    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as w:
        consultar_df('SELECT * FROM usuarios').to_excel(w, sheet_name='Usuarios', index=False)
        consultar_df('SELECT * FROM rotinas').to_excel(w, sheet_name='Rotinas', index=False)
        consultar_df('SELECT * FROM tarefas_dia').to_excel(w, sheet_name='Tarefas_Dia', index=False)
        consultar_df('SELECT * FROM projetos').to_excel(w, sheet_name='Projetos', index=False)
        consultar_df('SELECT * FROM historico').to_excel(w, sheet_name='Historico', index=False)
