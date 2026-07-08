# -*- coding: utf-8 -*-
from pathlib import Path
import pandas as pd
from database import init_db, conectar, criar_usuario, DB_PATH

def col(row,*names):
    for n in names:
        if n in row and pd.notna(row[n]): return str(row[n]).strip()
    return ''

def migrar_excel_inicial(excel_path='Agenda.xlsx', db_path=DB_PATH):
    p=Path(excel_path)
    if not p.exists(): return False, 'Agenda.xlsx não encontrado.'
    init_db(db_path); conn=conectar(db_path); xls=pd.ExcelFile(p); qu=qr=qp=0
    if 'Usuarios' in xls.sheet_names:
        df=pd.read_excel(p, sheet_name='Usuarios', dtype=object)
        for _,r in df.iterrows():
            nome=col(r,'Nome','nome')
            if nome: criar_usuario(nome, col(r,'Perfil') or 'Usuário', col(r,'Departamento')); qu+=1
    if 'Tarefas' in xls.sheet_names:
        df=pd.read_excel(p, sheet_name='Tarefas', dtype=object)
        for _,r in df.iterrows():
            rotina=col(r,'Tarefa','tarefa')
            if not rotina: continue
            conn.execute("""INSERT OR IGNORE INTO rotinas (rotina,descricao,departamento,responsavel,periodicidade,obrigatoria,prioridade,data_inicio,projeto,ativa,criado_por,criado_em,ultima_atualizacao) VALUES (?,?,?,?,?,?,?,?,?,?,'Migração','','')""", (rotina,col(r,'Descrição','Descricao','descricao'),col(r,'Departamento','departamento'),col(r,'Responsavel','Responsável','responsavel'),col(r,'Periodicidade','periodicidade') or 'Diaria',col(r,'Obrigatoria','Obrigatória','obrigatoria') or 'Não',col(r,'Prioridade','prioridade') or 'Normal',col(r,'Data de Inicio','Data Início','data_inicio'),col(r,'Projeto','projeto'),col(r,'Ativa','ativa') or 'Sim'))
            qr+=1
    if 'Projetos' in xls.sheet_names:
        df=pd.read_excel(p, sheet_name='Projetos', dtype=object)
        for _,r in df.iterrows():
            projeto=col(r,'Projeto','projeto')
            if not projeto: continue
            conn.execute("""INSERT OR IGNORE INTO projetos (projeto,objetivo,departamento,responsavel,data_inicio,prazo_final,status,proxima_etapa,observacao,ativo,criado_em) VALUES (?,?,?,?,?,?,?,?,?,?, '')""", (projeto,col(r,'Objetivo'),col(r,'Departamento'),col(r,'Responsavel','Responsável'),col(r,'Data de Inicio','Data Início'),col(r,'Prazo Final','prazo_final'),col(r,'Status') or 'Em andamento',col(r,'Proxima Etapa','Próxima Etapa'),col(r,'Observação','Observacao'),col(r,'Ativo') or 'Sim'))
            qp+=1
    conn.commit(); conn.close(); return True, f'Migração concluída: {qu} usuários, {qr} rotinas, {qp} projetos.'
if __name__ == '__main__': print(migrar_excel_inicial())
