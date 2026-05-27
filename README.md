# Agenda Operacional V5

Correção principal:
- Resolve erro: Invalid value 'Concluída' for dtype ...
- O app agora lê o Excel com dtype=object e força colunas em object antes de alterar Status.

Como usar:
1. Coloque Agenda.xlsx na mesma pasta do app.py.
2. Rode: pip install -r requirements.txt
3. Rode: streamlit run app.py


V6 Planner:
- Ao concluir, a tarefa sai das listas abertas.
- Concluídas hoje aparecem em seção própria.
- Tarefas concluídas aparecem riscadas, estilo planner.
- Departamentos ocultam concluídas por padrão, com opção para exibir.


V7:
- Corrige seletor de usuário no menu lateral.
- Remove frase explicativa abaixo do usuário.
- Ajusta horário para America/Sao_Paulo.


V8:
- Inclui tela "Editar tarefas".
- Permite alterar tarefas existentes.
- Permite arquivar tarefas sem excluir histórico.
- Registra alteração/arquivamento no Histórico.


V9:
- Restaura scroll vertical da tela principal.
- Restaura scroll da sidebar.
- Adiciona espaçamento inferior para telas longas.
- Inclui link "Voltar ao topo".
