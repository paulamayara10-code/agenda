# FIRST OPS 2.0 — Sprint 1 com Migrador

Sistema novo, com arquitetura limpa:

- Usuários
- Rotinas mestre
- Execuções diárias
- Pendências reais
- Projetos
- Histórico
- Administração
- Exportação

## Como usar

1. Mantenha `Agenda.xlsx` na mesma pasta do app.
2. Rode:

```bash
pip install -r requirements.txt
streamlit run app.py
```

3. Abra a aba **Migração**.
4. Clique em **IMPORTAR BACKUP PARA O FIRST OPS**.
5. Depois vá para **Home** e clique em **Gerar checklist do dia**.

## Observação

O Excel é usado apenas como backup inicial. A base oficial passa a ser o SQLite `first_ops.db`.
