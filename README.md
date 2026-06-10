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


V18.4 Restaurada e Segura:
- Base agenda.db restaurada: 44 tarefas, 3 projetos, 5 usuários e 28 históricos.
- Preservadas as 6 tarefas criadas por Claudia em 02/06.
- Preservadas as 4 conclusões de Claudia em 03/06.
- Removidas duplicidades causadas por reimportação.
- Migração forçada passa a rodar apenas se o banco estiver vazio, evitando duplicação.


V18.5 Data e Dias Úteis:
- Inclui seletor de Data de Referência no Dashboard, Coordenação, Minhas tarefas, Departamentos e Pendências com observação.
- Tarefas diárias aparecem apenas em dias úteis.
- Sábado/domingo não geram agenda recorrente.
- Tarefa única em fim de semana é empurrada para o próximo dia útil.


V18.6 Data Global:
- Inclui uma data global no topo de todas as abas.
- Botões: Dia anterior, Hoje e Próximo dia.
- Todas as abas passam a usar a mesma data de referência.
- Histórico filtra pela data selecionada, com opção de mostrar completo.
- Calendário usa a data global.


V18.7 Pendências anteriores:
- Corrige contador de pendências anteriores.
- Tarefa diária não concluída no último dia útil anterior passa a contar como pendência anterior.
- Fim de semana não gera pendência.
- Admin SQLite mostra diagnóstico de pendências anteriores calculadas.
