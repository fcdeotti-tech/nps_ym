import pandas as pd
import numpy as np
import warnings
import os
from datetime import datetime

warnings.filterwarnings('ignore')

# 1. DEFINIÇÃO DOS CAMINHOS
caminho_pasta = r"C:\Users\fernando.deotti\OneDrive - Holding Specialist Researchers Brasil LTDA\Documentos\Projetos\Yamaha\NPS\Python"
arquivo_pv = os.path.join(caminho_pasta, 'PV.xlsx')
arquivo_ve = os.path.join(caminho_pasta, 'VE.xlsx')

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
arquivo_saida = os.path.join(caminho_pasta, f'Analise_Detratores_Yamaha_{timestamp}.xlsx')

# 2. CARREGAR E LIMPAR DADOS BÁSICOS
print("Carregando bases para análise de Detratores...")
df_pv = pd.read_excel(arquivo_pv, sheet_name='PV', skiprows=8).dropna(how='all')
df_ve = pd.read_excel(arquivo_ve, sheet_name='VE', skiprows=8).dropna(how='all')

if 'Unnamed: 0' in df_pv.columns: df_pv = df_pv.drop(columns=['Unnamed: 0'])
if 'Unnamed: 0' in df_ve.columns: df_ve = df_ve.drop(columns=['Unnamed: 0'])

for col in ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']:
    for df_temp in [df_pv, df_ve]:
        if col in df_temp.columns:
            df_temp[col] = df_temp[col].replace('-', 'Não especificada').fillna('Não especificada')
            df_temp[col] = df_temp[col].apply(lambda x: 'Não especificada' if str(x).strip() == '' else x)

for col in ['Grupo', 'Região', 'Modelo']:
    if col in df_pv.columns: df_pv[col] = df_pv[col].fillna('Não informado')
    if col in df_ve.columns: df_ve[col] = df_ve[col].fillna('Não informado')

col_nps = 'Nota Recomendação concessionária (RESULTADO OFICIAL)'
cols_auxiliares = ['Nota Satisfação Atendimento', 'Recomendação da Marca', 'Recomendação da Moto Marca']

for df_temp in [df_pv, df_ve]:
    df_temp[col_nps] = pd.to_numeric(df_temp[col_nps], errors='coerce')
    for col_aux in cols_auxiliares:
        if col_aux in df_temp.columns:
            df_temp[col_aux] = pd.to_numeric(df_temp[col_aux], errors='coerce')

# 3. REGRA DE CLASSIFICAÇÃO: PROMOTOR, NEUTRO, DETRATOR +, DETRATOR -
def classificar_segmento_detrator(row):
    nota = row.get(col_nps)
    if pd.isna(nota): return np.nan
    
    if nota >= 9: return 'Promotor'
    elif nota >= 7: return 'Neutro'
    else: # Nota 0 a 6
        notas_validas = [row.get(c) for c in cols_auxiliares if c in row and pd.notna(row.get(c))]
        # Se a média das outras avaliações for muito alta (9 ou 10), ele é um Detrator+ (Culpa exclusiva da loja)
        if notas_validas and np.mean(notas_validas) >= 9:
            return 'Detrator +'
        else:
            return 'Detrator -'

df_ve['Segmento_NPS'] = df_ve.apply(classificar_segmento_detrator, axis=1)
df_pv['Segmento_NPS'] = df_pv.apply(classificar_segmento_detrator, axis=1)

def classificar_servico(tipo):
    if pd.isna(tipo): return 'Outros'
    return 'Venda de Peças' if 'PÇS' in str(tipo).upper() or 'OUTROS' in str(tipo).upper() else 'Revisão'
df_pv['Categoria_PV'] = df_pv['Tipo de Entrevista'].apply(classificar_servico)

# 4. FUNÇÃO PARA CALCULAR O NPS POTENCIAL (REVERTENDO DETRATOR+)
def calcular_potencial_detrator(df, dimensoes_agrupamento):
    resultados = []
    
    for nome_grupo, df_grupo in df.groupby(dimensoes_agrupamento):
        total = len(df_grupo['Segmento_NPS'].dropna())
        if total == 0: continue
            
        contagem = df_grupo['Segmento_NPS'].value_counts()
        prom = (contagem.get('Promotor', 0) / total) * 100
        neu = (contagem.get('Neutro', 0) / total) * 100
        detr_p = (contagem.get('Detrator +', 0) / total) * 100
        detr_m = (contagem.get('Detrator -', 0) / total) * 100
        
        # O NPS Atual
        nps_atual = prom - (detr_p + detr_m)
        
        # NPS Potencial: A loja não erra, o Detrator+ acompanha a nota da marca/moto e vira Promotor
        nps_potencial = (prom + detr_p) - detr_m
        ganho = nps_potencial - nps_atual
        
        linha = {d: v for d, v in zip(dimensoes_agrupamento, nome_grupo if isinstance(nome_grupo, tuple) else [nome_grupo])}
        linha.update({
            'Volume (N)': total,
            '% Detrator -': round(detr_m, 1),
            '% Detrator +': round(detr_p, 1),
            '% Neutro': round(neu, 1),
            '% Promotor': round(prom, 1),
            'NPS Atual': round(nps_atual, 1),
            'NPS Potencial': round(nps_potencial, 1),
            'Ganho Possível': round(ganho, 1)
        })
        resultados.append(linha)
        
    df_resultado = pd.DataFrame(resultados)
    if not df_resultado.empty:
        df_resultado = df_resultado.sort_values('Ganho Possível', ascending=False)
    return df_resultado

# 5. FUNÇÃO PARA ANALISAR CAUSAS DOS DETRATORES
def analisar_causas_detratores(df, dimensoes):
    df_detratores = df[df['Segmento_NPS'].isin(['Detrator +', 'Detrator -'])]
    if df_detratores.empty: return pd.DataFrame()
    
    agrupado = df_detratores.groupby(['Segmento_NPS'] + dimensoes).size().reset_index(name='Volume')
    totais_segmento = df_detratores['Segmento_NPS'].value_counts()
    agrupado['% no Segmento'] = agrupado.apply(lambda row: round((row['Volume'] / totais_segmento[row['Segmento_NPS']]) * 100, 1), axis=1)
    
    return agrupado.sort_values(by=['Segmento_NPS', 'Volume'], ascending=[False, False])

# 6. GERAR EXCEL
print("Gerando análises de Detratores...")
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    
    # --- 1. VISÃO NPS POTENCIAL ---
    calcular_potencial_detrator(df_ve, ['Região']).to_excel(writer, sheet_name='VE_Potencial_Regiao', index=False)
    calcular_potencial_detrator(df_ve, ['Grupo']).to_excel(writer, sheet_name='VE_Potencial_Grupo', index=False)
    calcular_potencial_detrator(df_ve, ['Concessionária']).to_excel(writer, sheet_name='VE_Potencial_Conc', index=False)
    
    calcular_potencial_detrator(df_pv, ['Região', 'Categoria_PV']).to_excel(writer, sheet_name='PV_Potencial_Regiao', index=False)
    calcular_potencial_detrator(df_pv, ['Grupo', 'Categoria_PV']).to_excel(writer, sheet_name='PV_Potencial_Grupo', index=False)
    calcular_potencial_detrator(df_pv, ['Concessionária', 'Categoria_PV']).to_excel(writer, sheet_name='PV_Potencial_Conc', index=False)
    
    # --- 2. VISÃO CAUSAS DOS DETRATORES ---
    analisar_causas_detratores(df_ve, ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']).to_excel(writer, sheet_name='VE_Causas_Detrat', index=False)
    analisar_causas_detratores(df_pv, ['Categoria_PV', 'Causa da nota de recomendação', 'Subcausa da nota de recomendação']).to_excel(writer, sheet_name='PV_Causas_Detrat', index=False)
    
    # --- 3. DISTRIBUIÇÃO POR MODELO ---
    analisar_causas_detratores(df_ve, ['Modelo']).to_excel(writer, sheet_name='VE_Modelos_Detrat', index=False)
    
print(f"\nArquivo de Análise de Detratores gerado com sucesso!\n{arquivo_saida}")