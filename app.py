# -*- coding: utf-8 -*-
from datetime import datetime, date, timedelta
from pathlib import Path
import unicodedata
import pandas as pd
import streamlit as st
from database import *
from migrar_excel_inicial import migrar_excel_inicial

st.set_page_config(page_title='Agenda Operacional V20', page_icon='✅', layout='wide', initial_sidebar_state='expanded')
init_db(DB_PATH)
if Path('Agenda.xlsx').exists() and consultar_df('SELECT COUNT(*) qtd FROM rotinas').iloc[0]['qtd']==0:
    ok_migracao, msg_migracao=migrar_excel_inicial('Agenda.xlsx', DB_PATH)
else: ok_migracao, msg_migracao=False,''

st.markdown('''<style>.stApp{background:#f5f7fb}div.block-container{padding-top:1.3rem;padding-bottom:4rem;max-width:1500px}section[data-testid="stSidebar"]{background:linear-gradient(180deg,#0f172a,#111827 70%,#1e1b4b)}section[data-testid="stSidebar"] *{color:#fff!important}section[data-testid="stSidebar"] div[data-baseweb="select"] *{color:#0f172a!important}.title{font-size:38px;font-weight:900;color:#0f172a;letter-spacing:-.04em}.subtitle{color:#64748b;margin-bottom:18px}.panel{background:white;border:1px solid #e2e8f0;border-radius:24px;padding:20px;box-shadow:0 14px 35px rgba(15,23,42,.06);margin-top:16px}.panel-title{font-size:22px;font-weight:900;color:#0f172a;margin-bottom:14px}.metric-card{background:white;border:1px solid #e2e8f0;border-radius:22px;padding:20px;box-shadow:0 12px 30px rgba(15,23,42,.06);min-height:120px}.metric-label{font-size:12px;color:#64748b;font-weight:800;text-transform:uppercase}.metric-value{font-size:36px;font-weight:900;margin-top:8px}.tag{display:inline-block;padding:6px 10px;border-radius:999px;font-size:12px;font-weight:800;margin-right:6px;margin-top:6px}.blue{background:#dbeafe;color:#1d4ed8}.green{background:#dcfce7;color:#15803d}.red{background:#fee2e2;color:#b91c1c}.yellow{background:#fef3c7;color:#b45309}.purple{background:#ede9fe;color:#6d28d9}.gray{background:#f1f5f9;color:#475569}.orange{background:#ffedd5;color:#c2410c}.task-title{font-size:18px;font-weight:900;color:#0f172a}.done{text-decoration:line-through;color:#94a3b8}.stButton>button{border-radius:14px;font-weight:800}</style>''', unsafe_allow_html=True)

def txt(v): return '' if pd.isna(v) else str(v).strip()
def norm(v): return ' '.join(unicodedata.normalize('NFKD', txt(v).lower()).encode('ascii','ignore').decode('ascii').split())
def hoje(): return date.today()
def br(d): return d.strftime('%d/%m/%Y')
def parse_br(v):
    try: return pd.to_datetime(v, dayfirst=True).date() if txt(v) else None
    except Exception: return None
def eh_dia_util(d): return d.weekday()<5
def data_global():
    if 'data_ref' not in st.session_state: st.session_state['data_ref']=hoje()
    return st.session_state['data_ref']
def mudar_data(dias): st.session_state['data_ref']=data_global()+timedelta(days=dias)
def voltar_hoje(): st.session_state['data_ref']=hoje()
def normalizar_departamento(v):
    mapa={'contas a receber':'Contas a Receber','contas a pagar':'Contas a Pagar','contabilidade':'Contabilidade','controladoria':'Controladoria','tesouraria':'Tesouraria','financeiro':'Financeiro','projetos':'Projetos'}
    return mapa.get(norm(v), txt(v).title())
def periodicidade(v):
    t=norm(v)
    if t in ['', 'diario','diaria','todo dia','todos os dias']: return 'diaria'
    if t in ['semanal','semana']: return 'semanal'
    if t in ['mensal','mes']: return 'mensal'
    if t in ['unica','unico','pontual']: return 'unica'
    return t
def rotina_deve_gerar(row, data_ref):
    if not eh_dia_util(data_ref): return False
    inicio=parse_br(row.get('data_inicio'))
    if inicio and inicio>data_ref: return False
    per=periodicidade(row.get('periodicidade'))
    if per=='diaria': return True
    if per=='semanal': return (data_ref-(inicio or data_ref)).days%7==0
    if per=='mensal': return data_ref.day==(inicio or data_ref).day
    if per=='unica': return inicio==data_ref
    return True
def gerar_tarefas_do_dia(data_ref):
    rotinas=listar_rotinas(True); qtd=0; dt=br(data_ref)
    for _,r in rotinas.iterrows():
        if rotina_deve_gerar(r, data_ref): criar_tarefa_dia(r, dt); qtd+=1
    return qtd

def metric_card(label,value,detail,color): st.markdown(f"<div class='metric-card'><div class='metric-label'>{label}</div><div class='metric-value' style='color:{color}'>{value}</div><div style='color:{color};font-size:13px'>{detail}</div></div>", unsafe_allow_html=True)
def header(t,s): st.markdown(f"<div class='title'>{t}</div><div class='subtitle'>{s}</div>", unsafe_allow_html=True)
def data_bar():
    st.markdown("<div class='panel' style='padding:14px 18px'>", unsafe_allow_html=True); c1,c2,c3,c4,c5=st.columns([1.3,1.2,.9,.7,2.2])
    with c1: st.markdown('#### 📅 Data operacional')
    with c2: st.session_state['data_ref']=st.date_input('Data', value=data_global(), label_visibility='collapsed')
    with c3: st.button('◀ Anterior', use_container_width=True, on_click=mudar_data, args=(-1,))
    with c4: st.button('Hoje', use_container_width=True, on_click=voltar_hoje)
    with c5:
        a,b=st.columns([.8,1.5])
        with a: st.button('Próximo ▶', use_container_width=True, on_click=mudar_data, args=(1,))
        with b: st.success(f"Dia útil: {br(data_global())}") if eh_dia_util(data_global()) else st.warning('Fim de semana: não gera rotina diária.')
    st.markdown('</div>', unsafe_allow_html=True); return data_global()
def frame(t,s): header(t,s); return data_bar()
def sidebar():
    st.sidebar.markdown('## ✅ Agenda'); st.sidebar.markdown('### Operacional V20'); st.sidebar.divider(); users=listar_usuarios(); nomes=users['nome'].dropna().tolist() if not users.empty else []
    if not nomes: criar_usuario('Paula','Administradora',''); nomes=['Paula']
    user=st.sidebar.selectbox('Usuário', nomes); deps=[normalizar_departamento(d) for d in listar_departamentos()]; deps=sorted(list(dict.fromkeys([d for d in deps if d])))
    menu=['Painel do Dia','Coordenação','Minhas tarefas']+deps+['Pendências anteriores','Rotinas','Projetos','Histórico','Administração','Exportar']
    return user, st.sidebar.radio('Navegação', menu), deps
def status_tag(s): return {'Pendente':'yellow','Em andamento':'orange','Concluída':'green','Cancelada':'gray','Reprogramada':'blue'}.get(s,'purple')
def task_card(row,user,prefix):
    status=txt(row.get('status')) or 'Pendente'; tid=int(row.get('id'))
    with st.container(border=True):
        c1,c2=st.columns([4.7,1.45])
        with c1:
            cls='task-title done' if status=='Concluída' else 'task-title'; st.markdown(f"<div class='{cls}'>{txt(row.get('tarefa'))}</div>", unsafe_allow_html=True)
            if txt(row.get('descricao')): st.caption(txt(row.get('descricao')))
            tags=f"<span class='tag {status_tag(status)}'>{status}</span>"
            for val,css,prefixo in [(row.get('departamento'),'blue',''),(row.get('responsavel'),'green','👤 '),(row.get('projeto'),'purple','📁 '),(row.get('prioridade'),'yellow','')]:
                if txt(val): tags+=f"<span class='tag {css}'>{prefixo}{txt(val)}</span>"
            st.markdown(tags, unsafe_allow_html=True)
            if txt(row.get('observacao')): st.info(txt(row.get('observacao')))
        with c2:
            if status!='Concluída':
                if status!='Em andamento' and st.button('▶ Iniciar', key=f'ini_{prefix}_{tid}'): atualizar_status_tarefa(tid,'Em andamento',user,'Iniciada'); st.rerun()
                if st.button('✅ Concluir', key=f'conc_{prefix}_{tid}'): atualizar_status_tarefa(tid,'Concluída',user,'Concluída'); st.rerun()
            else: st.success('Concluída')
        with st.expander('💬 Justificar / reprogramar / cancelar'):
            obs=st.text_area('Observação', key=f'obs_{prefix}_{tid}')
            A,B,C=st.columns(3)
            with A:
                if st.button('Salvar observação', key=f'save_{prefix}_{tid}'): adicionar_observacao_tarefa(tid,user,obs); st.rerun()
            with B:
                nd=st.date_input('Reprogramar para', value=data_global()+timedelta(days=1), key=f'repdate_{prefix}_{tid}')
                if st.button('Reprogramar', key=f'rep_{prefix}_{tid}'): reprogramar_tarefa(tid,user,br(nd),obs or 'Reprogramada'); st.rerun()
            with C:
                if st.button('Cancelar', key=f'can_{prefix}_{tid}'): cancelar_tarefa(tid,user,obs or 'Cancelada'); st.rerun()

def painel_dia(user):
    ref=frame('Painel do Dia','Checklist operacional gerado por data')
    if ok_migracao and msg_migracao: st.success(msg_migracao)
    if st.button('🔄 Gerar/atualizar checklist do dia', type='primary'): gerar_tarefas_do_dia(ref); st.success(f'Checklist atualizado para {br(ref)}.')
    tarefas=listar_tarefas_dia(br(ref)); pend=listar_pendencias_anteriores(br(ref)); abertas=tarefas[~tarefas['status'].isin(['Concluída','Cancelada','Reprogramada'])] if not tarefas.empty else tarefas; concl=tarefas[tarefas['status']=='Concluída'] if not tarefas.empty else tarefas; minhas=abertas[abertas['responsavel'].astype(str).str.lower().str.contains(user.lower(),na=False)] if not abertas.empty else abertas
    c1,c2,c3,c4=st.columns(4)
    with c1: metric_card('Tarefas do dia',len(tarefas),br(ref),'#2563eb')
    with c2: metric_card('Abertas',len(abertas),'Pendentes do dia','#f59e0b')
    with c3: metric_card('Pendências anteriores',len(pend),'Dias passados','#dc2626')
    with c4: metric_card('Concluídas',len(concl),'Na data','#16a34a')
    if tarefas.empty and eh_dia_util(ref): st.info("Clique em 'Gerar/atualizar checklist do dia' para criar as tarefas da data selecionada.")
    st.markdown("<div class='panel'><div class='panel-title'>⭐ Minhas tarefas</div>", unsafe_allow_html=True)
    if minhas.empty: st.success('Sem tarefas pendentes para você nesta data.')
    else:
        for _,r in minhas.iterrows(): task_card(r,user,'minhas')
    st.markdown('</div>', unsafe_allow_html=True)
    A,B=st.columns(2)
    with A:
        st.markdown("<div class='panel'><div class='panel-title'>🟡 Checklist do dia</div>", unsafe_allow_html=True)
        if abertas.empty: st.success('Checklist do dia finalizado.')
        else:
            for _,r in abertas.iterrows(): task_card(r,user,'dia')
        st.markdown('</div>', unsafe_allow_html=True)
    with B:
        st.markdown("<div class='panel'><div class='panel-title'>🔴 Pendências anteriores</div>", unsafe_allow_html=True)
        if pend.empty: st.success('Sem pendências anteriores.')
        else:
            for _,r in pend.iterrows(): task_card(r,user,'pend')
        st.markdown('</div>', unsafe_allow_html=True)

def coordenacao(user):
    ref=frame('Coordenação','Gestão da equipe e gargalos operacionais'); tarefas=listar_tarefas_dia(br(ref)); pend=listar_pendencias_anteriores(br(ref)); abertas=tarefas[~tarefas['status'].isin(['Concluída','Cancelada','Reprogramada'])] if not tarefas.empty else tarefas; andam=abertas[abertas['status']=='Em andamento'] if not abertas.empty else abertas; concl=tarefas[tarefas['status']=='Concluída'] if not tarefas.empty else tarefas
    c1,c2,c3,c4,c5=st.columns(5)
    with c1: metric_card('Do dia',len(tarefas),br(ref),'#2563eb')
    with c2: metric_card('Abertas',len(abertas),'Executar','#f59e0b')
    with c3: metric_card('Em andamento',len(andam),'Iniciadas','#f97316')
    with c4: metric_card('Pend. anteriores',len(pend),'Gargalos','#dc2626')
    with c5: metric_card('Concluídas',len(concl),'Na data','#16a34a')
    st.markdown("<div class='panel'><div class='panel-title'>📊 Ranking operacional do dia</div>", unsafe_allow_html=True)
    if tarefas.empty: st.info('Sem tarefas geradas para esta data.')
    else:
        rows=[]
        for resp in sorted([x for x in tarefas['responsavel'].dropna().unique() if txt(x)]):
            b=tarefas[tarefas['responsavel'].astype(str)==resp]; rows.append({'Usuário':resp,'Total':len(b),'Abertas':len(b[~b['status'].isin(['Concluída','Cancelada','Reprogramada'])]),'Em andamento':len(b[b['status']=='Em andamento']),'Concluídas':len(b[b['status']=='Concluída'])})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

def minhas_tarefas(user):
    ref=frame('Minhas tarefas',f'Checklist de {user}'); tarefas=listar_tarefas_dia(br(ref)); minhas=tarefas[tarefas['responsavel'].astype(str).str.lower().str.contains(user.lower(),na=False)] if not tarefas.empty else tarefas; abertas=minhas[~minhas['status'].isin(['Concluída','Cancelada','Reprogramada'])] if not minhas.empty else minhas
    if abertas.empty: st.success('Sem tarefas pendentes para você.')
    else:
        for _,r in abertas.iterrows(): task_card(r,user,'minhas_page')

def departamento_page(user,depto):
    ref=frame(depto,f'Checklist do departamento {depto}'); tarefas=listar_tarefas_dia(br(ref)); base=tarefas[tarefas['departamento'].astype(str).str.lower()==depto.lower()] if not tarefas.empty else tarefas
    if base.empty: st.info('Sem tarefas para este departamento na data selecionada.'); return
    if not st.toggle('Mostrar finalizadas', value=False): base=base[~base['status'].isin(['Concluída','Cancelada','Reprogramada'])]
    for _,r in base.iterrows(): task_card(r,user,f'dep_{depto}')

def pendencias_anteriores_page(user):
    ref=frame('Pendências anteriores','Execuções passadas não concluídas'); pend=listar_pendencias_anteriores(br(ref))
    if pend.empty: st.success('Sem pendências anteriores.')
    else:
        for _,r in pend.iterrows(): task_card(r,user,'pend_page')

def rotinas(user):
    ref=frame('Rotinas','Cadastro mestre das atividades recorrentes')
    with st.expander('➕ Nova rotina', expanded=False):
        deps=listar_departamentos() or ['Contas a Receber','Contas a Pagar','Contabilidade','Controladoria','Tesouraria']; users=listar_usuarios(); resps=['']+(users['nome'].tolist() if not users.empty else []); projs=listar_projetos(); projs_l=['']+(projs['projeto'].tolist() if not projs.empty else [])
        with st.form('nova_rotina'):
            a,b=st.columns(2)
            with a: rotina=st.text_input('Rotina'); desc=st.text_area('Descrição'); dep=st.selectbox('Departamento',deps); resp=st.selectbox('Responsável',resps); proj=st.selectbox('Projeto',projs_l)
            with b: per=st.selectbox('Periodicidade',['Diaria','Semanal','Mensal','Unica']); obrig=st.selectbox('Obrigatória',['Não','Sim']); prio=st.selectbox('Prioridade',['Normal','Alta','Crítica','Baixa']); inicio=st.date_input('Data de início', value=ref)
            salvar=st.form_submit_button('Salvar rotina')
        if salvar: criar_rotina({'rotina':rotina,'descricao':desc,'departamento':normalizar_departamento(dep),'responsavel':resp,'periodicidade':per,'obrigatoria':obrig,'prioridade':prio,'data_inicio':br(inicio),'projeto':proj},user); st.success('Rotina criada.'); st.rerun()
    df=listar_rotinas(True); st.dataframe(df, use_container_width=True, hide_index=True) if not df.empty else st.info('Nenhuma rotina cadastrada.')

def projetos(user):
    ref=frame('Projetos','Gestão separada dos projetos')
    with st.expander('➕ Novo projeto', expanded=False):
        users=listar_usuarios(); resps=['']+(users['nome'].tolist() if not users.empty else [])
        with st.form('novo_projeto'):
            projeto=st.text_input('Projeto'); objetivo=st.text_area('Objetivo'); dep=st.text_input('Departamento'); resp=st.selectbox('Responsável',resps); prazo=st.date_input('Prazo final', value=ref); etapa=st.text_input('Próxima etapa'); salvar=st.form_submit_button('Salvar projeto')
        if salvar: criar_projeto({'projeto':projeto,'objetivo':objetivo,'departamento':dep,'responsavel':resp,'prazo_final':br(prazo),'proxima_etapa':etapa},user); st.success('Projeto criado.'); st.rerun()
    projs=listar_projetos(); tarefas=listar_tarefas_dia(incluir_todas=True); resumo=[]
    for _,p in projs.iterrows():
        nome=txt(p.get('projeto')); vinc=tarefas[tarefas['projeto'].astype(str).str.lower()==nome.lower()] if not tarefas.empty else pd.DataFrame(); total=len(vinc); concl=len(vinc[vinc['status']=='Concluída']) if total else 0; pct=round((concl/total)*100,1) if total else 0; resumo.append({'Projeto':nome,'Responsável':p.get('responsavel'),'Prazo':p.get('prazo_final'),'Progresso %':pct,'Tarefas':total,'Concluídas':concl,'Status':'Concluído' if pct==100 and total else p.get('status'),'Próxima etapa':p.get('proxima_etapa')})
    st.dataframe(pd.DataFrame(resumo), use_container_width=True, hide_index=True) if resumo else st.info('Nenhum projeto cadastrado.')

def historico():
    ref=frame('Histórico','Auditoria das movimentações'); hist=listar_historico()
    if hist.empty: st.info('Sem histórico.'); return
    if not st.toggle('Mostrar tudo', value=False): hist=hist[hist['data_ref'].astype(str)==br(ref)]
    st.dataframe(hist, use_container_width=True, hide_index=True)

def administracao(user):
    ref=frame('Administração','Base, usuários e recuperação'); c1,c2,c3=st.columns(3)
    with c1: st.metric('Rotinas',len(listar_rotinas(False)))
    with c2: st.metric('Tarefas geradas',len(listar_tarefas_dia(incluir_todas=True)))
    with c3: st.metric('Usuários',len(listar_usuarios()))
    with st.form('novo_user'):
        nome=st.text_input('Nome'); perfil=st.selectbox('Perfil',['Usuário','Administradora']); dep=st.text_input('Departamento'); salvar=st.form_submit_button('Criar usuário')
    if salvar and nome: criar_usuario(nome,perfil,dep); st.success('Usuário criado.'); st.rerun()
    st.dataframe(listar_usuarios(), use_container_width=True, hide_index=True)

def exportar(user):
    ref=frame('Exportar','Baixar base completa'); saida=Path('Agenda_V20_Exportada.xlsx')
    if st.button('Gerar Excel'): exportar_excel(saida); st.success('Exportação gerada.')
    if saida.exists(): st.download_button('Baixar Excel', data=saida.read_bytes(), file_name=saida.name)

def main():
    user,page,deps=sidebar()
    if page=='Painel do Dia': painel_dia(user)
    elif page=='Coordenação': coordenacao(user)
    elif page=='Minhas tarefas': minhas_tarefas(user)
    elif page in deps: departamento_page(user,page)
    elif page=='Pendências anteriores': pendencias_anteriores_page(user)
    elif page=='Rotinas': rotinas(user)
    elif page=='Projetos': projetos(user)
    elif page=='Histórico': historico()
    elif page=='Administração': administracao(user)
    elif page=='Exportar': exportar(user)
if __name__=='__main__': main()
