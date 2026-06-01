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

V10:
- Corrige conclusão por data: tarefa diária concluída hoje não aparece concluída amanhã.
- Calendário considera corretamente a data selecionada.
- Normaliza departamentos para evitar menu duplicado.


V11 Coordenação:
- Nova aba Coordenação.
- Dashboard gerencial com atrasadas, críticas, abertas, concluídas e sem responsável.
- Painel de gargalos.
- Ranking operacional da equipe.
- Kanban operacional.
- Feed de atividades com base no Histórico.
- Backup automático antes de salvar alterações.


V12 UI/UX:
- Redesign visual mais forte.
- Aba Coordenação garantida no menu.
- Kanban com 4 colunas.
- Cards mais modernos.
- Sidebar mais profissional.
- Painel executivo para apresentação à equipe.

V13 Premium:
- Oculta barra superior padrão do Streamlit.
- Sidebar mais premium.
- Hero no Dashboard.
- Cards e métricas com visual mais moderno.
- Data/hora em pill no topo.


V14 Sidebar fixa:
- Sidebar fixada à esquerda.
- Remove botão de recolher sidebar quando aberta.
- Botão de sidebar recolhida fica visível caso o navegador preserve estado recolhido.
- Conteúdo principal é deslocado para não ficar atrás do menu.


V15:
- Corrige sobreposição da sidebar fixa.
- Reduz largura da sidebar para 240px.
- Desloca área principal corretamente para a direita.
- Ajusta largura útil dos cards e painéis.


V16 Observações e Segurança:
- Campo de observações/justificativa dentro de cada tarefa.
- Observações ficam registradas no Histórico.
- Lock simples de gravação para reduzir conflito com múltiplos usuários.
- Backup automático preservado antes de salvar.
- Tela "Tarefas arquivadas" para consultar e reativar tarefas.
- Tarefas não são excluídas pelo sistema; apenas arquivadas.
