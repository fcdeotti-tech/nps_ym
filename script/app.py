import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
import os

# Configuração da Página
st.set_page_config(page_title="Yamaha NPS Dashboard", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 1. FUNÇÕES DE CARREGAMENTO E CAMINHOS
# ==========================================

# Identifica a pasta onde este script (app.py) está salvo
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@st.cache_data
def load_data():
    # Define os caminhos subindo um nível e entrando na pasta /source
    caminho_pv = os.path.join(BASE_DIR, '..', 'source', 'PV.xlsx')
    caminho_ve = os.path.join(BASE_DIR, '..', 'source', 'VE.xlsx')
    
    # Verifica se os arquivos existem antes de tentar ler
    if not os.path.exists(caminho_pv) or not os.path.exists(caminho_ve):
        st.error(f"Erro: Arquivos não encontrados na pasta source! \n\nProcurado em: {os.path.join(BASE_DIR, '..', 'source')}")
        st.stop()

    # Leitura dos arquivos
    df_pv = pd.read_excel(caminho_pv, sheet_name='PV', skiprows=8).dropna(how='all')
    df_ve = pd.read_excel(caminho_ve, sheet_name='VE', skiprows=8).dropna(how='all')
    
    # Limpeza e Padronização
    for df in [df_pv, df_ve]:
        if 'Unnamed: 0' in df.columns: df.drop(columns=['Unnamed: 0'], inplace=True)
        for col in ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']:
            if col in df.columns:
                df[col] = df[col].replace('-', 'Não especificada').fillna('Não especificada')
                df[col] = df[col].apply(lambda x: 'Não especificada' if str(x).strip() == '' else x)
        for col in ['Grupo', 'Região', 'Modelo', 'Concessionária']:
            if col in df.columns: df[col] = df[col].fillna('Não informado')
            
        # Converter Notas
        col_nps_orig = 'Nota Recomendação concessionária (RESULTADO OFICIAL)'
        df[col_nps_orig] = pd.to_numeric(df[col_nps_orig], errors='coerce')
        
        for col_aux in ['Nota Satisfação Atendimento', 'Recomendação da Marca', 'Recomendação da Moto Marca']:
            if col_aux in df.columns: df[col_aux] = pd.to_numeric(df[col_aux], errors='coerce')
            
    # Classificação Avançada (Neutros e Detratores)
    def classificar_segmento(row):
        nota = row.get('Nota Recomendação concessionária (RESULTADO OFICIAL)')
        if pd.isna(nota): return np.nan
        if nota >= 9: return 'Promotor'
        
        notas_aux = [row.get(c) for c in ['Nota Satisfação Atendimento', 'Recomendação da Marca', 'Recomendação da Moto Marca'] if pd.notna(row.get(c))]
        media_aux = np.mean(notas_aux) if notas_aux else 0
        
        if nota >= 7: return 'Neutro +' if media_aux >= 9 else 'Neutro -'
        else: return 'Detrator +' if media_aux >= 9 else 'Detrator -'

    df_ve['Segmento_Av'] = df_ve.apply(classificar_segmento, axis=1)
    df_pv['Segmento_Av'] = df_pv.apply(classificar_segmento, axis=1)
    
    # Extrair Revisões no PV
    def cat_revisao(texto):
        if pd.isna(texto): return "Não é Revisão"
        nums = re.findall(r'\d+', str(texto))
        if not nums: return "Não é Revisão"
        return f"Revisão {int(nums[0])}" if int(nums[0]) <= 4 else "Revisão 5 ou +"
        
    df_pv['Ciclo_Revisao'] = df_pv['Tipo de Entrevista'].apply(cat_revisao)
    
    return df_ve, df_pv

# Executa a carga
df_ve, df_pv = load_data()

# ==========================================
# 2. BARRA LATERAL E FILTROS DINÂMICOS
# ==========================================
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Yamaha_Motor_Logo.svg/1024px-Yamaha_Motor_Logo.svg.png", width=150)
st.sidebar.title("Filtros Yamaha")

departamento = st.sidebar.radio("Escolha o Departamento", ['Vendas (VE)', 'Pós-Vendas (PV)'])
df_base = df_ve if departamento == 'Vendas (VE)' else df_pv

# Filtros que se adaptam à base selecionada
regioes = st.sidebar.multiselect("Filtrar por Região", sorted(df_base['Região'].unique()))
if regioes: df_base = df_base[df_base['Região'].isin(regioes)]

grupos = st.sidebar.multiselect("Filtrar por Grupo", sorted(df_base['Grupo'].unique()))
if grupos: df_base = df_base[df_base['Grupo'].isin(grupos)]

concessionarias = st.sidebar.multiselect("Filtrar por Concessionária", sorted(df_base['Concessionária'].unique()))
if concessionarias: df_base = df_base[df_base['Concessionária'].isin(concessionarias)]

modelos = st.sidebar.multiselect("Filtrar por Modelo", sorted(df_base['Modelo'].unique()))
if modelos: df_base = df_base[df_base['Modelo'].isin(modelos)]

col_nps = 'Nota Recomendação concessionária (RESULTADO OFICIAL)'

# ==========================================
# 3. FUNÇÕES AUXILIARES DE CÁLCULO
# ==========================================
def calc_nps_from_df(df):
    notas = df[col_nps].dropna()
    if len(notas) == 0: return 0, 0
    prom = len(notas[notas >= 9])
    detr = len(notas[notas <= 6])
    return ((prom - detr) / len(notas)) * 100, len(notas)

def calc_drivers(df):
    nps_total, n_total = calc_nps_from_df(df)
    if n_total == 0: return pd.DataFrame()
    
    dados = []
    for causa, grupo in df.groupby('Causa da nota de recomendação'):
        if causa in ['Outro(s) motivo(s)', 'Não especificada']: continue
        nps_causa, n_causa = calc_nps_from_df(grupo)
        perc_coluna = n_causa / n_total
        contrib = perc_coluna * nps_causa
        peso = (contrib / nps_total) * 100 if nps_total != 0 else 0
        gap = peso - (perc_coluna * 100)
        dados.append({
            'Causa': causa, 'Volume (N)': n_causa, 'NPS': round(nps_causa, 2), 
            'Contribuição': round(contrib, 2), 'Gap': round(gap, 2)
        })
    return pd.DataFrame(dados).sort_values('Gap', ascending=True)

# ==========================================
# 4. CONSTRUÇÃO DAS ABAS
# ==========================================
st.title(f"Dashboard NPS: {departamento}")

tabs = st.tabs([
    "📊 Visão Geral", "🎯 Gaps por Causa", 
    "⚖️ Análise Neutros", "🚨 Análise Detratores", 
    "🔧 Revisões", "💬 Verbatims", 
    "📝 Resumo Executivo"
])

# ABA 1: VISÃO GERAL
with tabs[0]:
    nps_geral, n_geral = calc_nps_from_df(df_base)
    c1, c2, c3 = st.columns(3)
    c1.metric("NPS Atual", f"{nps_geral:.1f}")
    c2.metric("Total Entrevistas", f"{n_geral}")
    
    prom = len(df_base[df_base[col_nps] >= 9])
    detr = len(df_base[df_base[col_nps] <= 6])
    neu = len(df_base[(df_base[col_nps] >= 7) & (df_base[col_nps] <= 8)])
    c3.metric("P | N | D", f"{prom} | {neu} | {detr}")

    st.markdown("---")
    col_a, col_b = st.columns(2)
    
    df_reg = df_base.groupby('Região').apply(lambda x: calc_nps_from_df(x)[0]).reset_index(name='NPS').sort_values('NPS', ascending=False)
    fig_reg = px.bar(df_reg, x='Região', y='NPS', title="NPS por Região", color='NPS', color_continuous_scale='Blues')
    col_a.plotly_chart(fig_reg, use_container_width=True)
    
    df_mod = df_base.groupby('Modelo').apply(lambda x: calc_nps_from_df(x)[0]).reset_index(name='NPS')
    df_mod = df_mod[df_mod['Modelo'] != 'Não informado'].sort_values('NPS', ascending=False).head(10)
    fig_mod = px.bar(df_mod, x='Modelo', y='NPS', title="Top 10 Modelos", color='NPS', color_continuous_scale='Blues')
    col_b.plotly_chart(fig_mod, use_container_width=True)

# ABA 2: CAUSAS E GAPS
with tabs[1]:
    st.subheader("Direcionadores da Nota (Drivers)")
    df_drivers = calc_drivers(df_base)
    if not df_drivers.empty:
        fig_gap = px.bar(df_drivers, y='Causa', x='Gap', orientation='h', color='Gap', color_continuous_scale='RdYlGn', text_auto=True)
        st.plotly_chart(fig_gap, use_container_width=True)
        st.dataframe(df_drivers, use_container_width=True)
    else:
        st.info("Filtros muito restritos para gerar a análise de causas.")

# ABA 3: NEUTROS
with tabs[2]:
    neu_p = len(df_base[df_base['Segmento_Av'] == 'Neutro +'])
    nps_atual = nps_geral
    prom_n = len(df_base[df_base['Segmento_Av'] == 'Promotor'])
    detr_n = len(df_base[df_base['Segmento_Av'].str.contains('Detrator', na=False)])
    
    nps_potencial = (((prom_n + neu_p) - detr_n) / n_geral * 100) if n_geral > 0 else 0
    
    c1, c2 = st.columns(2)
    c1.metric("NPS Atual", f"{nps_atual:.1f}")
    c2.metric("NPS Potencial (Neutro+ -> Promotor)", f"{nps_potencial:.1f}", f"+{nps_potencial - nps_atual:.1f}")
    
    st.write(f"Existem **{neu_p}** clientes classificados como 'Neutro +'.")
    df_neu_motivos = df_base[df_base['Segmento_Av'] == 'Neutro +']['Causa da nota de recomendação'].value_counts()
    st.bar_chart(df_neu_motivos)

# ABA 4: DETRATORES
with tabs[3]:
    detr_plus = len(df_base[df_base['Segmento_Av'] == 'Detrator +'])
    st.write(f"Temos **{detr_plus}** Detratores que amam o produto/marca mas odiaram a loja (Detratores+).")
    df_det_motivos = df_base[df_base['Segmento_Av'] == 'Detrator +']['Causa da nota de recomendação'].value_counts()
    st.bar_chart(df_det_motivos)

# ABA 5: REVISÕES (SÓ PARA PV)
with tabs[4]:
    if departamento == 'Pós-Vendas (PV)':
        df_rev = df_base[df_base['Ciclo_Revisao'] != 'Não é Revisão']
        rev_nps = df_rev.groupby('Ciclo_Revisao').apply(lambda x: calc_nps_from_df(x)[0]).reset_index(name='NPS')
        fig_rev = px.line(rev_nps, x='Ciclo_Revisao', y='NPS', markers=True, title="Desempenho por Estágio de Revisão")
        st.plotly_chart(fig_rev, use_container_width=True)
    else:
        st.info("Selecione 'Pós-Vendas' no menu lateral para ver esta análise.")

# ABA 6: VERBATIMS
with tabs[5]:
    df_verb = df_base.dropna(subset=['Comentários'])
    if not df_verb.empty:
        sel_causa = st.selectbox("Escolha uma Causa para ler os comentários", ['Todas'] + list(df_verb['Causa da nota de recomendação'].unique()))
        if sel_causa != 'Todas': df_verb = df_verb[df_verb['Causa da nota de recomendação'] == sel_causa]
        st.dataframe(df_verb[['Concessionária', 'Modelo', 'Nota Recomendação concessionária (RESULTADO OFICIAL)', 'Causa da nota de recomendação', 'Comentários']])
    else:
        st.write("Sem comentários para os filtros aplicados.")

# ABA 7: EXECUTIVE SUMMARY
with tabs[6]:
    st.markdown("""
    ### 📌 Resumo Estratégico
    1. **Foco no Detrator+:** Clientes que gostam da Yamaha mas tiveram problemas na loja. Prioridade total em **Entrega Técnica** e **Agilidade**.
    2. **Conversão de Neutros+:** Ganho imediato de NPS ao melhorar processos de financiamento e qualidade básica de oficina.
    3. **Alerta 3ª Revisão:** Monitorar o choque de preço para evitar a perda do cliente para o mercado paralelo.
    """)