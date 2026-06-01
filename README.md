# Agenda Operacional V18.1 SQLite Auto

## Como usar

1. Coloque `Agenda.xlsx` na mesma pasta do `app.py`.
2. Suba/rode o app normalmente:

```bash
streamlit run app.py
```

Na primeira execução, se o `agenda.db` estiver vazio, o app importa automaticamente o `Agenda.xlsx`.

## Importante

- O SQLite (`agenda.db`) vira a base principal.
- O Excel passa a ser usado para migração inicial e exportação.
- Se precisar forçar nova migração, use o menu **Admin SQLite**.

## Arquivos

- app.py
- database.py
- migrar_excel_para_sqlite.py
- requirements.txt
