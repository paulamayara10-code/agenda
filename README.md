# Agenda Operacional V18 SQLite

## Como usar

1. Coloque `Agenda.xlsx` na mesma pasta, se quiser migrar sua base antiga.
2. Instale dependências:

```bash
pip install -r requirements.txt
```

3. Migre o Excel para SQLite uma única vez:

```bash
python migrar_excel_para_sqlite.py
```

4. Rode o app:

```bash
streamlit run app.py
```

## Arquivos

- app.py: aplicativo Streamlit
- database.py: camada SQLite
- migrar_excel_para_sqlite.py: migração do Excel para agenda.db
- agenda.db: banco criado automaticamente
- requirements.txt

## Observação

O SQLite passa a ser a base principal. O Excel fica como importação/exportação.
