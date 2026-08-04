# FIRST OPS

Agenda operacional em Streamlit para gestão de rotinas, tarefas, equipe, projetos, histórico e backups.

## Arquivos principais

- `app.py`: aplicação Streamlit.
- `first_ops_database_v232.py`: banco SQLite e regras operacionais.
- `migrate.py`: carga inicial a partir do `Agenda.xlsx`.
- `Agenda.xlsx`: cadastros iniciais de usuários, tarefas, projetos e histórico.
- `requirements.txt`: dependências Python.
- `runtime.txt`: versão do Python no Streamlit Cloud.

## Publicação no Streamlit Cloud

- Branch: `main`
- Main file path: `app.py`

O banco é criado em `data/first_ops_enterprise2.db`. A pasta `data/` e os backups locais não devem ser versionados no GitHub.

Antes de substituir uma implantação existente, baixe pelo app uma cópia do banco SQLite e uma exportação Excel.
