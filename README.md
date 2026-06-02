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


V18.2 Projetos:
- Projetos com progresso automático.
- Projeto é concluído automaticamente quando 100% das tarefas vinculadas estiverem concluídas.
- Cadastro de novos projetos pelo app.
- Cards de projetos com tarefas vinculadas.
- Comentários de projeto registrados no histórico.
- Painel de projetos na Coordenação.


V18.3 Periodicidade:
- Corrige leitura de tarefas diárias.
- Reconhece Diario, Diário, Diaria, Diária e variações.
- Tarefas diárias aparecem todos os dias enquanto não forem concluídas no dia.
- Tarefas sem data de início não somem.
- Admin SQLite mostra diagnóstico de periodicidade.
