# FIRST OPS Enterprise

Nesta versão, as tarefas são calculadas diretamente a partir das rotinas.

- Abrir uma data não grava dezenas de execuções no banco.
- A execução é criada somente ao iniciar, concluir, comentar, reprogramar ou cancelar.
- Rotinas diárias, semanais, mensais e únicas aparecem automaticamente na data correta.
- Pendências anteriores usam o último dia útil.
- Base e Backup, Administração e Exportação ficam restritos ao perfil administrador.

## Executar

```bash
pip install -r requirements.txt
streamlit run app.py
```


## FIRST OPS Enterprise 1.1

Correção técnica:
- O módulo `db.py` foi renomeado para `first_ops_database.py`.
- Todos os imports foram atualizados.
- A alteração evita conflito com módulos externos chamados `db` no Streamlit Cloud.


## FIRST OPS Enterprise 1.2

Correções:
- Responsáveis, departamentos e periodicidades passam a reconhecer os nomes
  das colunas em minúsculas do backup.
- A opção "Corrigir cadastros importados" completa as 44 rotinas já existentes,
  sem apagar histórico ou execuções.
- Tarefas compartilhadas continuam reconhecendo responsáveis separados por `/`.
- O nome selecionado no campo Usuário passou a ter contraste forçado.
