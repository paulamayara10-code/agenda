# FIRST OPS Enterprise 2.0

Versão reconstruída com banco novo e isolado.

## Dados carregados

- 5 usuários
- 44 rotinas
- 3 projetos
- 28 registros históricos legados

## Recursos

- atividades distribuídas pelo campo `responsavel` do backup;
- tarefas compartilhadas reconhecidas por `/`;
- geração diária automática;
- pendências de dias anteriores;
- iniciar, concluir, comentar, reprogramar e cancelar;
- histórico diário de todas as ações;
- backup automático local por dia;
- download do banco SQLite e exportação Excel;
- Base e Backup disponível apenas para administradores.

## Publicação

Coloque todos os arquivos e pastas do ZIP na raiz do repositório e execute:

```bash
streamlit run app.py
```

O banco desta versão é `data/first_ops_enterprise2.db`, portanto não usa bancos antigos.


## FIRST OPS Enterprise 2.1 — Go Live

- Início oficial da operação: **10/07/2026**.
- Pendências anteriores zeradas até essa data.
- Novas pendências começam a ser contabilizadas no próximo dia útil.
- Sábados e domingos não geram atividades.
- O backfill nunca cria atividades anteriores ao Go Live.
- Incluído indicador de atividades com prazo/horário no dia.


## FIRST OPS Enterprise 2.2 — Praticidade e Backup

- A tela inicial agora é individual.
- Totais gerais permanecem nas telas Equipe e Coordenação.
- Checklist rápido para concluir atividades.
- Detalhes, observações, reprogramação e cancelamento ficam recolhidos.
- Backup automático diário.
- Backup completo manual em um clique.
- Mantidas as 30 cópias mais recentes.
- Download direto do Excel e do SQLite.


## FIRST OPS Enterprise 2.2.1 — Ambiente Estável

Correção de publicação no Streamlit Cloud:

- Streamlit fixado em `1.48.1`.
- pandas fixado em `2.3.3`.
- NumPy fixado em `2.3.2`.
- openpyxl fixado em `3.1.5`.
- Python fixado na linha `3.11`.
- Evita que o ambiente instale automaticamente versões futuras incompatíveis.


## FIRST OPS Enterprise 2.2.2 — Equipe Estável

Correções:
- A tela Equipe não utiliza mais `st.dataframe`.
- Os colaboradores são exibidos em cards leves.
- É possível abrir as atividades de cada pessoa individualmente.
- PyArrow fixado em `20.0.0` para evitar falhas nativas no Streamlit Cloud.
