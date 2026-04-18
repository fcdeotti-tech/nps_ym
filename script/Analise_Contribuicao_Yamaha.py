import pandas as pd
import numpy as np
import warnings
import os
from datetime import datetime

warnings.filterwarnings('ignore')

# 1. DEFINIÇÃO DOS CAMINHOS (WSL / LINUX)
DIRETORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
DIRETORIO_RAIZ = os.path.dirname(DIRETORIO_SCRIPT)
DIRETORIO_SOURCE = os.path.join(DIRETORIO_RAIZ, 'source')
DIRETORIO_OUTPUT = os.path.join(DIRETORIO_RAIZ, 'output')

arquivo_pv = os.path.join(DIRETORIO_SOURCE, 'PV.xlsx')
arquivo_ve = os.path.join(DIRETORIO_SOURCE, 'VE.xlsx')
arquivo_saida = os.path.join(DIRETORIO_OUTPUT, 'Analise_NPS_Yamaha.xlsx')

# 2. FUNÇÃO DE CÁLCULO DE NPS
def calcular_nps(notas):
    notas = pd.to_numeric(notas, errors='coerce').dropna()
    if len(notas) == 0: return np.nan
    return ((len(notas[notas >= 9]) - len(notas[notas <= 6])) / len(notas)) * 100

# 3. FUNÇÃO MESTRE DE CONTRIBUIÇÃO (COM HIERARQUIA)
def gerar_analise_contribuicao(df, col_nps, cols_hierarquia):
    total_n = len(df[col_nps].dropna())
    total_nps = calcular_nps(df[col_nps])
    
    # Agrupamento com Causa e Subcausa
    cols_agrupamento = cols_hierarquia + ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']
    
    res = df.groupby(cols_agrupamento).agg(
        N_valido=(col_nps, lambda x: x.notna().sum()),
        NPS=(col_nps, calcular_nps)
    ).reset_index()
    
    # Cálculos de Contribuição e Impacto (Gap)
    res['Contribuição'] = (res['N_valido'] / total_n) * res['NPS']
    res['Gap'] = ((res['Contribuição'] / total_nps) * 100) - ((res['N_valido'] / total_n) * 100) if total_nps != 0 else 0
    res['N_%_da_coluna'] = (res['N_valido'] / total_n) * 100
    
    return res

# 4. PROCESSAMENTO
print("Processando dados e injetando hierarquia (Região/Grupo)...")
df_ve = pd.read_excel(arquivo_ve, sheet_name='VE', skiprows=8).dropna(how='all')
df_pv = pd.read_excel(arquivo_pv, sheet_name='PV', skiprows=8).dropna(how='all')

col_nps = 'Nota Recomendação concessionária (RESULTADO OFICIAL)'

with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    # Vendas - Níveis
    gerar_analise_contribuicao(df_ve, col_nps, ['Região']).to_excel(writer, sheet_name='VE_Reg_C_Sub', index=False)
    gerar_analise_contribuicao(df_ve, col_nps, ['Região', 'Grupo']).to_excel(writer, sheet_name='VE_Grup_C_Sub', index=False)
    gerar_analise_contribuicao(df_ve, col_nps, ['Região', 'Grupo', 'Concessionária']).to_excel(writer, sheet_name='VE_Conc_C_Sub', index=False)
    gerar_analise_contribuicao(df_ve, col_nps, ['Modelo']).to_excel(writer, sheet_name='VE_Mod_C_Sub', index=False)
    
    # Totais Nacionais
    df_ve_total = gerar_analise_contribuicao(df_ve, col_nps, [])
    df_ve_total.to_excel(writer, sheet_name='VE_Tot_C_Sub', index=False)
    df_ve_total.groupby('Causa da nota de recomendação').agg({'N_valido':'sum', 'NPS':'mean', 'Contribuição':'sum', 'Gap':'sum'}).reset_index().to_excel(writer, sheet_name='VE_Tot_Causa', index=False)

    # Pós-Vendas - Total (Soma de Revisão e Venda de Peças)
    gerar_analise_contribuicao(df_pv, col_nps, ['Região']).to_excel(writer, sheet_name='PV_Tot_Reg_C_Sub', index=False)
    gerar_analise_contribuicao(df_pv, col_nps, ['Região', 'Grupo']).to_excel(writer, sheet_name='PV_Tot_Grup_C_Sub', index=False)
    gerar_analise_contribuicao(df_pv, col_nps, ['Região', 'Grupo', 'Concessionária']).to_excel(writer, sheet_name='PV_Tot_Conc_C_Sub', index=False)
    
    df_pv_total = gerar_analise_contribuicao(df_pv, col_nps, [])
    df_pv_total.to_excel(writer, sheet_name='PV_Tot_C_Sub', index=False)
    df_pv_total.groupby('Causa da nota de recomendação').agg({'N_valido':'sum', 'NPS':'mean', 'Contribuição':'sum', 'Gap':'sum'}).reset_index().to_excel(writer, sheet_name='PV_Tot_Causa', index=False)

print(f"Arquivo consolidado com hierarquia gerado em: {arquivo_saida}")