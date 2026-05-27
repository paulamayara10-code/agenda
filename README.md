# Agenda Operacional V5

Correção principal:
- Resolve erro: Invalid value 'Concluída' for dtype ...
- O app agora lê o Excel com dtype=object e força colunas em object antes de alterar Status.

Como usar:
1. Coloque Agenda.xlsx na mesma pasta do app.py.
2. Rode: pip install -r requirements.txt
3. Rode: streamlit run app.py
