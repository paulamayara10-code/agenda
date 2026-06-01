
# -*- coding: utf-8 -*-
from pathlib import Path
import pandas as pd

from database import init_db, conectar, criar_usuario, DB_PATH


EXCEL_PADRAO = "Agenda.xlsx"


def col(row, *nomes):
    for nome in nomes:
        if nome in row and pd.notna(row[nome]):
            return str(row[nome]).strip()
    return ""


def migrar(excel_path=EXCEL_PADRAO, db_path=DB_PATH):
    excel_path = Path(excel_path)

    if not excel_path.exists():
        print(f"Arquivo não encontrado: {excel_path}")
        return

    init_db(db_path)
    conn = conectar(db_path)

    xls = pd.ExcelFile(excel_path)

    if "Usuarios" in xls.sheet_names:
        usuarios = pd.read_excel(excel_path, sheet_name="Usuarios", dtype=object)
        for _, row in usuarios.iterrows():
            nome = col(row, "Nome", "nome")
            perfil = col(row, "Perfil", "perfil") or "Usuário"
            departamento = col(row, "Departamento", "departamento")
            criar_usuario(nome, perfil, departamento, db_path)

    if "Projetos" in xls.sheet_names:
        projetos = pd.read_excel(excel_path, sheet_name="Projetos", dtype=object)
        for _, row in projetos.iterrows():
            conn.execute(
                """
                INSERT INTO projetos (
                    projeto, objetivo, departamento, responsavel, data_inicio, prazo_final,
                    status, percentual, proxima_etapa, observacao, ativo
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    col(row, "Projeto"),
                    col(row, "Objetivo"),
                    col(row, "Departamento"),
                    col(row, "Responsavel", "Responsável"),
                    col(row, "Data de Inicio", "Data Início"),
                    col(row, "Prazo Final", "Prazo"),
                    col(row, "Status") or "Em andamento",
                    float(col(row, "%", "Percentual") or 0),
                    col(row, "Proxima Etapa", "Próxima Etapa"),
                    col(row, "Observação", "Observacao"),
                    col(row, "Ativo") or "Sim",
                )
            )

    if "Tarefas" in xls.sheet_names:
        tarefas = pd.read_excel(excel_path, sheet_name="Tarefas", dtype=object)
        for _, row in tarefas.iterrows():
            tarefa = col(row, "Tarefa")
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
                    col(row, "Descrição", "Descricao"),
                    col(row, "Departamento"),
                    col(row, "Projeto"),
                    col(row, "Responsavel", "Responsável"),
                    col(row, "Periodicidade"),
                    col(row, "Obrigatoria", "Obrigatória") or "Não",
                    col(row, "Prioridade") or "Normal",
                    col(row, "Dependencia", "Dependência"),
                    col(row, "Prazo Limite"),
                    col(row, "Data de Inicio", "Data Início") or "",
                    col(row, "Status") or "Pendente",
                    col(row, "Concluído Por", "Concluido Por"),
                    col(row, "Data Conclusão", "Data Conclusao"),
                    col(row, "Observação", "Observacao"),
                    col(row, "Ativa") or "Sim",
                    "Migração",
                    "",
                    "",
                )
            )

    if "Historico" in xls.sheet_names:
        hist = pd.read_excel(excel_path, sheet_name="Historico", dtype=object)
        for _, row in hist.iterrows():
            conn.execute(
                """
                INSERT INTO historico (id_tarefa, tarefa, usuario, data, hora, status, observacao, criado_em)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    col(row, "ID Tarefa") or None,
                    col(row, "Tarefa"),
                    col(row, "Usuário", "Usuario"),
                    col(row, "Data"),
                    col(row, "Hora"),
                    col(row, "Status"),
                    col(row, "Observação", "Observacao"),
                    "",
                )
            )

    conn.commit()
    conn.close()
    print(f"Migração concluída para {db_path}")


if __name__ == "__main__":
    migrar()
