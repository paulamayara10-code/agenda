
# -*- coding: utf-8 -*-
from pathlib import Path
import pandas as pd
from database import init_db, conectar, criar_usuario, DB_PATH, banco_tem_dados


def col(row, *nomes):
    for nome in nomes:
        if nome in row and pd.notna(row[nome]):
            return str(row[nome]).strip()
    return ""


def to_float(v):
    try:
        t = str(v).replace("%", "").replace(",", ".").strip()
        return float(t) if t else 0
    except Exception:
        return 0


def migrar_excel_para_sqlite(excel_path="Agenda.xlsx", db_path=DB_PATH, somente_se_vazio=True):
    excel_path = Path(excel_path)

    if not excel_path.exists():
        return False, f"Arquivo {excel_path} não encontrado."

    init_db(db_path)

    if somente_se_vazio and banco_tem_dados(db_path):
        return False, "Banco já possui dados. Migração automática não executada."

    conn = conectar(db_path)
    xls = pd.ExcelFile(excel_path)

    total_usuarios = 0
    total_tarefas = 0
    total_projetos = 0
    total_historico = 0

    if "Usuarios" in xls.sheet_names:
        usuarios = pd.read_excel(excel_path, sheet_name="Usuarios", dtype=object)
        for _, row in usuarios.iterrows():
            nome = col(row, "Nome", "nome")
            if not nome:
                continue
            perfil = col(row, "Perfil", "perfil") or "Usuário"
            departamento = col(row, "Departamento", "departamento")
            criar_usuario(nome, perfil, departamento, db_path)
            total_usuarios += 1

    if "Projetos" in xls.sheet_names:
        projetos = pd.read_excel(excel_path, sheet_name="Projetos", dtype=object)
        for _, row in projetos.iterrows():
            projeto = col(row, "Projeto", "projeto")
            if not projeto:
                continue
            conn.execute(
                """
                INSERT INTO projetos (
                    projeto, objetivo, departamento, responsavel, data_inicio, prazo_final,
                    status, percentual, proxima_etapa, observacao, ativo
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    projeto,
                    col(row, "Objetivo", "objetivo"),
                    col(row, "Departamento", "departamento"),
                    col(row, "Responsavel", "Responsável", "responsavel"),
                    col(row, "Data de Inicio", "Data Início", "data_inicio"),
                    col(row, "Prazo Final", "Prazo", "prazo_final"),
                    col(row, "Status", "status") or "Em andamento",
                    to_float(col(row, "%", "Percentual", "percentual")),
                    col(row, "Proxima Etapa", "Próxima Etapa", "proxima_etapa"),
                    col(row, "Observação", "Observacao", "observacao"),
                    col(row, "Ativo", "ativo") or "Sim",
                )
            )
            total_projetos += 1

    if "Tarefas" in xls.sheet_names:
        tarefas = pd.read_excel(excel_path, sheet_name="Tarefas", dtype=object)
        for _, row in tarefas.iterrows():
            tarefa = col(row, "Tarefa", "tarefa")
            if not tarefa:
                continue

            conn.execute(
                """
                INSERT INTO tarefas (
                    tarefa, descricao, departamento, projeto, responsavel, periodicidade,
                    obrigatoria, prioridade, dependencia, prazo_limite, data_inicio,
                    status, concluido_por, data_conclusao, observacao, ativa,
                    criado_por, data_criacao, ultima_atualizacao
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    tarefa,
                    col(row, "Descrição", "Descricao", "descricao"),
                    col(row, "Departamento", "departamento"),
                    col(row, "Projeto", "projeto"),
                    col(row, "Responsavel", "Responsável", "responsavel"),
                    col(row, "Periodicidade", "periodicidade") or "Diario",
                    col(row, "Obrigatoria", "Obrigatória", "obrigatoria") or "Não",
                    col(row, "Prioridade", "prioridade") or "Normal",
                    col(row, "Dependencia", "Dependência", "dependencia"),
                    col(row, "Prazo Limite", "prazo_limite"),
                    col(row, "Data de Inicio", "Data Início", "data_inicio"),
                    col(row, "Status", "status") or "Pendente",
                    col(row, "Concluído Por", "Concluido Por", "concluido_por"),
                    col(row, "Data Conclusão", "Data Conclusao", "data_conclusao"),
                    col(row, "Observação", "Observacao", "observacao"),
                    col(row, "Ativa", "ativa") or "Sim",
                    "Migração Excel",
                    "",
                    "",
                )
            )
            total_tarefas += 1

    if "Historico" in xls.sheet_names:
        hist = pd.read_excel(excel_path, sheet_name="Historico", dtype=object)
        for _, row in hist.iterrows():
            conn.execute(
                """
                INSERT INTO historico (id_tarefa, tarefa, usuario, data, hora, status, observacao, criado_em)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    col(row, "ID Tarefa", "id_tarefa") or None,
                    col(row, "Tarefa", "tarefa"),
                    col(row, "Usuário", "Usuario", "usuario"),
                    col(row, "Data", "data"),
                    col(row, "Hora", "hora"),
                    col(row, "Status", "status"),
                    col(row, "Observação", "Observacao", "observacao"),
                    "",
                )
            )
            total_historico += 1

    conn.commit()
    conn.close()

    return True, f"Migração concluída: {total_usuarios} usuários, {total_tarefas} tarefas, {total_projetos} projetos, {total_historico} históricos."


if __name__ == "__main__":
    ok, msg = migrar_excel_para_sqlite("Agenda.xlsx", DB_PATH, somente_se_vazio=False)
    print(msg)
