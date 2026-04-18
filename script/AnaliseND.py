import pandas as pd
import numpy as np
import warnings
import os
from datetime import datetime

warnings.filterwarnings('ignore')

# ==========================================
# 1. DEFINIÇÃO DOS CAMINHOS (WSL / LINUX)
# ==========================================
DIRETORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
DIRETORIO_RAIZ = os.path.dirname(DIRETORIO_SCRIPT)

DIRETORIO_SOURCE = os.path.join(DIRETORIO_RAIZ, 'source')
DIRETORIO_OUTPUT = os.path.join(DIRETORIO_RAIZ, 'output')

arquivo_pv = os.path.join(DIRETORIO_SOURCE, 'PV.xlsx')
arquivo_ve = os.path.join(DIRETORIO_SOURCE, 'VE.xlsx')

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
arq_saida_neutros = os.path.join(DIRETORIO_OUTPUT, f'Analise_Neutros_Yamaha_{timestamp}.xlsx')
arq_saida_detratores = os.path.join(DIRETORIO_OUTPUT, f'Analise_Detratores_Yamaha_{timestamp}.xlsx')

# ==========================================
# 2. CARREGAR E LIMPAR DADOS BÁSICOS
# ==========================================
print("Carregando bases de dados (VE e PV)...")
df_pv = pd.read_excel(arquivo_pv, sheet_name='PV', skiprows=8).dropna(how='all')
df_ve = pd.read_excel(arquivo_ve, sheet_name='VE', skiprows=8).dropna(how='all')

if 'Unnamed: 0' in df_pv.columns: df_pv = df_pv.drop(columns=['Unnamed: 0'])
if 'Unnamed: 0' in df_ve.columns: df_ve = df_ve.drop(columns=['Unnamed: 0'])

# Tratamento de textos vazios
for col in ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']:
    for df_temp in [df_pv, df_ve]:
        if col in df_temp.columns:
            df_temp[col] = df_temp[col].replace('-', 'Não especificada').fillna('Não especificada')
            df_temp[col] = df_temp[col].apply(lambda x: 'Não especificada' if str(x).strip() == '' else x)

for col in ['Grupo', 'Região', 'Modelo', 'Concessionária']:
    if col in df_pv.columns: df_pv[col] = df_pv[col].fillna('Não informado')
    if col in df_ve.columns: df_ve[col] = df_ve[col].fillna('Não informado')

# Converter Notas para números
col_nps = 'Nota Recomendação concessionária (RESULTADO OFICIAL)'
cols_auxiliares = ['Nota Satisfação Atendimento', 'Recomendação da Marca', 'Recomendação da Moto Marca']

for df_temp in [df_pv, df_ve]:
    df_temp[col_nps] = pd.to_numeric(df_temp[col_nps], errors='coerce')
    for col_aux in cols_auxiliares:
        if col_aux in df_temp.columns:
            df_temp[col_aux] = pd.to_numeric(df_temp[col_aux], errors='coerce')

# ==========================================
# 3. REGRA DE CLASSIFICAÇÃO UNIFICADA (5 SEGMENTOS)
# ==========================================
def classificar_segmento_unificado(row):
    nota = row.get(col_nps)
    if pd.isna(nota): return np.nan
    
    notas_validas = [row.get(c) for c in cols_auxiliares if c in row and pd.notna(row.get(c))]
    media_aux = np.mean(notas_validas) if notas_validas else 0
    
    if nota >= 9:
        return 'Promotor'
    elif nota >= 7:
        return 'Neutro +' if media_aux >= 9 else 'Neutro -'
    else:
        return 'Detrator +' if media_aux >= 9 else 'Detrator -'

df_ve['Segmento_NPS'] = df_ve.apply(classificar_segmento_unificado, axis=1)
df_pv['Segmento_NPS'] = df_pv.apply(classificar_segmento_unificado, axis=1)

def classificar_servico(tipo):
    if pd.isna(tipo): return 'Outros'
    return 'Venda de Peças' if 'PÇS' in str(tipo).upper() or 'OUTROS' in str(tipo).upper() else 'Revisão'
df_pv['Categoria_PV'] = df_pv['Tipo de Entrevista'].apply(classificar_servico)

# ==========================================
# 4. CÁLCULO DE NPS POTENCIAL (Com Totais Injetados)
# ==========================================
def calcular_potencial(df, dimensoes_agrupamento, foco='Neutros'):
    df_exp = df.copy()
    
    # Injetar Total Nacional
    if 'Região' in dimensoes_agrupamento:
        df_nac = df_exp.copy()
        df_nac['Região'] = 'TOTAL NACIONAL'
        df_exp = pd.concat([df_exp, df_nac], ignore_index=True)
        
    # Injetar Total PV
    if 'Categoria_PV' in dimensoes_agrupamento:
        df_cat = df_exp.copy()
        df_cat['Categoria_PV'] = 'TOTAL PV'
        df_exp = pd.concat([df_exp, df_cat], ignore_index=True)

    resultados = []
    for nome_grupo, df_grupo in df_exp.groupby(dimensoes_agrupamento):
        total = len(df_grupo['Segmento_NPS'].dropna())
        if total == 0: continue
            
        contagem = df_grupo['Segmento_NPS'].value_counts()
        prom = (contagem.get('Promotor', 0) / total) * 100
        n_p = (contagem.get('Neutro +', 0) / total) * 100
        n_m = (contagem.get('Neutro -', 0) / total) * 100
        d_p = (contagem.get('Detrator +', 0) / total) * 100
        d_m = (contagem.get('Detrator -', 0) / total) * 100
        
        nps_atual = prom - (d_p + d_m)
        
        # Lógica de conversão
        if foco == 'Neutros':
            nps_potencial = (prom + n_p) - (d_p + d_m)
        else: # Detratores
            nps_potencial = (prom + d_p) - d_m
            
        ganho = nps_potencial - nps_atual
        
        linha = {d: v for d, v in zip(dimensoes_agrupamento, nome_grupo if isinstance(nome_grupo, tuple) else [nome_grupo])}
        linha.update({
            'Volume (N)': total,
            '% Detrator -': round(d_m, 1),
            '% Detrator +': round(d_p, 1),
            '% Neutro -': round(n_m, 1),
            '% Neutro +': round(n_p, 1),
            '% Promotor': round(prom, 1),
            'NPS Atual': round(nps_atual, 1),
            'NPS Potencial': round(nps_potencial, 1),
            'Ganho Possível': round(ganho, 1)
        })
        resultados.append(linha)
        
    df_res = pd.DataFrame(resultados)
    
    # Ordenação Inteligente: Manter TOTAL NACIONAL e TOTAL PV no topo
    if not df_res.empty:
        sort_cols, sort_asc = [], []
        if 'Região' in df_res.columns:
            df_res['is_total_reg'] = df_res['Região'] == 'TOTAL NACIONAL'
            sort_cols.append('is_total_reg'); sort_asc.append(False)
        if 'Categoria_PV' in df_res.columns:
            df_res['is_total_cat'] = df_res['Categoria_PV'] == 'TOTAL PV'
            sort_cols.append('is_total_cat'); sort_asc.append(False)
            
        sort_cols.append('Ganho Possível')
        sort_asc.append(False)
        
        df_res = df_res.sort_values(by=sort_cols, ascending=sort_asc)
        df_res = df_res.drop(columns=[c for c in ['is_total_reg', 'is_total_cat'] if c in df_res.columns])
        
    return df_res

# ==========================================
# 5. ANÁLISE DE CAUSAS (Com Totais Injetados)
# ==========================================
def analisar_causas(df, dimensoes, foco='Neutros'):
    df_exp = df.copy()
    
    if 'Categoria_PV' in dimensoes:
        df_cat = df_exp.copy()
        df_cat['Categoria_PV'] = 'TOTAL PV'
        df_exp = pd.concat([df_exp, df_cat], ignore_index=True)

    segmentos_alvo = ['Neutro +', 'Neutro -'] if foco == 'Neutros' else ['Detrator +', 'Detrator -']
    df_filtro = df_exp[df_exp['Segmento_NPS'].isin(segmentos_alvo)]
    if df_filtro.empty: return pd.DataFrame()
    
    # Agrupamento base
    agrupado = df_filtro.groupby(['Segmento_NPS'] + dimensoes).size().reset_index(name='Volume')
    
    # Contexto para calcular o percentual corretamente (Ex: % dentro do TOTAL PV ou dentro de Revisão)
    context_dims = [d for d in dimensoes if d not in ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']]
    
    if context_dims:
        totais_contexto = df_filtro.groupby(['Segmento_NPS'] + context_dims).size().reset_index(name='Total_Segmento')
        agrupado = agrupado.merge(totais_contexto, on=['Segmento_NPS'] + context_dims)
    else:
        totais_contexto = df_filtro.groupby(['Segmento_NPS']).size().reset_index(name='Total_Segmento')
        agrupado = agrupado.merge(totais_contexto, on=['Segmento_NPS'])
        
    agrupado['% no Segmento'] = (agrupado['Volume'] / agrupado['Total_Segmento'] * 100).round(1)
    agrupado = agrupado.drop(columns=['Total_Segmento'])
    
    # Injetar o TOTAL DA CAUSA na Subcausa
    if 'Causa da nota de recomendação' in dimensoes and 'Subcausa da nota de recomendação' in dimensoes:
        dim_sem_sub = [d for d in dimensoes if d != 'Subcausa da nota de recomendação']
        agr_causa = df_filtro.groupby(['Segmento_NPS'] + dim_sem_sub).size().reset_index(name='Volume')
        
        if context_dims:
            agr_causa = agr_causa.merge(totais_contexto, on=['Segmento_NPS'] + context_dims)
        else:
            agr_causa = agr_causa.merge(totais_contexto, on=['Segmento_NPS'])
            
        agr_causa['% no Segmento'] = (agr_causa['Volume'] / agr_causa['Total_Segmento'] * 100).round(1)
        agr_causa = agr_causa.drop(columns=['Total_Segmento'])
        agr_causa['Subcausa da nota de recomendação'] = 'TOTAL DA CAUSA'
        
        agrupado = pd.concat([agrupado, agr_causa], ignore_index=True)

    # Ordenação Inteligente: 'TOTAL PV' no topo, seguido do 'TOTAL DA CAUSA', ordenando o resto por volume
    sort_cols, sort_asc = ['Segmento_NPS'], [False]
    
    if 'Categoria_PV' in agrupado.columns:
        agrupado['is_total_cat'] = agrupado['Categoria_PV'] == 'TOTAL PV'
        sort_cols.extend(['is_total_cat', 'Categoria_PV'])
        sort_asc.extend([False, True])
        
    if 'Causa da nota de recomendação' in dimensoes:
        agrupado['is_total_causa'] = agrupado['Subcausa da nota de recomendação'] == 'TOTAL DA CAUSA'
        sort_cols.extend(['Causa da nota de recomendação', 'is_total_causa'])
        sort_asc.extend([True, False])
        
    sort_cols.append('Volume')
    sort_asc.append(False)
    
    agrupado = agrupado.sort_values(by=sort_cols, ascending=sort_asc)
    agrupado = agrupado.drop(columns=[c for c in ['is_total_cat', 'is_total_causa'] if c in agrupado.columns])
    
    return agrupado

# ==========================================
# 6. GERAÇÃO DOS ARQUIVOS FINAIS
# ==========================================
def exportar_foco(ficheiro_saida, foco):
    print(f"Gerando relatórios e cálculos para: {foco}...")
    with pd.ExcelWriter(ficheiro_saida, engine='openpyxl') as writer:
        
        # Abas de Potencial (Concessionária, Grupo, Região)
        calcular_potencial(df_ve, ['Região'], foco).to_excel(writer, sheet_name='VE_Potencial_Regiao', index=False)
        calcular_potencial(df_ve, ['Grupo'], foco).to_excel(writer, sheet_name='VE_Potencial_Grupo', index=False)
        calcular_potencial(df_ve, ['Concessionária'], foco).to_excel(writer, sheet_name='VE_Potencial_Conc', index=False)
        
        calcular_potencial(df_pv, ['Região', 'Categoria_PV'], foco).to_excel(writer, sheet_name='PV_Potencial_Regiao', index=False)
        calcular_potencial(df_pv, ['Grupo', 'Categoria_PV'], foco).to_excel(writer, sheet_name='PV_Potencial_Grupo', index=False)
        calcular_potencial(df_pv, ['Concessionária', 'Categoria_PV'], foco).to_excel(writer, sheet_name='PV_Potencial_Conc', index=False)
        
        # Abas de Causas e Subcausas (Onde mora o problema)
        analisar_causas(df_ve, ['Causa da nota de recomendação', 'Subcausa da nota de recomendação'], foco).to_excel(writer, sheet_name=f'VE_Causas_{foco[:6]}', index=False)
        analisar_causas(df_pv, ['Categoria_PV', 'Causa da nota de recomendação', 'Subcausa da nota de recomendação'], foco).to_excel(writer, sheet_name=f'PV_Causas_{foco[:6]}', index=False)
        
        # Abas de Modelos de Motos
        analisar_causas(df_ve, ['Modelo'], foco).to_excel(writer, sheet_name=f'VE_Modelos_{foco[:6]}', index=False)

# Chamando as rotinas para gerar os dois arquivos
exportar_foco(arq_saida_neutros, 'Neutros')
exportar_foco(arq_saida_detratores, 'Detratores')

print(f"\nProcesso finalizado com sucesso! Arquivos salvos em:\n{DIRETORIO_OUTPUT}")