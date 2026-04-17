import pandas as pd
import numpy as np
import warnings
import os
import re
from datetime import datetime

warnings.filterwarnings('ignore')

# 1. DEFINIÇÃO DOS CAMINHOS
caminho_pasta = r"C:\Users\fernando.deotti\OneDrive - Holding Specialist Researchers Brasil LTDA\Documentos\Projetos\Yamaha\NPS\Python"
arquivo_pv = os.path.join(caminho_pasta, 'PV.xlsx')
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
arquivo_saida = os.path.join(caminho_pasta, f'Analise_Revisoes_Yamaha_{timestamp}.xlsx')

# 2. FUNÇÃO DE CÁLCULO DE NPS
def calcular_nps(notas):
    notas = pd.to_numeric(notas, errors='coerce').dropna()
    if len(notas) == 0: return np.nan
    promotores = len(notas[notas >= 9])
    detratores = len(notas[notas <= 6])
    return ((promotores - detratores) / len(notas)) * 100

# 3. FUNÇÃO ANALÍTICA MESTRE (Com Totais e Gaps)
def gerar_analise_contribuicao(df, colunas_agrupamento, dim_principal=None, col_nps='Nota Recomendação concessionária (RESULTADO OFICIAL)'):
    def processar_bloco(df_bloco, cols_group):
        total_n = len(df_bloco[col_nps].dropna())
        total_nps = calcular_nps(df_bloco[col_nps])
        if total_n == 0: return pd.DataFrame(), None
            
        agrupado = df_bloco.groupby(cols_group).agg(
            N_valido=(col_nps, lambda x: x.notna().sum()),
            NPS=(col_nps, calcular_nps)
        ).reset_index()
        
        agrupado = agrupado[agrupado['N_valido'] > 0]
        agrupado['N_%_da_coluna'] = (agrupado['N_valido'] / total_n) * 100
        
        contribuicao_real = (agrupado['N_valido'] / total_n) * agrupado['NPS']
        agrupado['Contribuição'] = contribuicao_real
        agrupado['Peso'] = (contribuicao_real / total_nps) * 100 if total_nps != 0 else 0
        agrupado['Gap'] = agrupado['Peso'] - agrupado['N_%_da_coluna']
        
        for col in ['N_%_da_coluna', 'Peso', 'Contribuição', 'Gap', 'NPS']:
            agrupado[col] = agrupado[col].round(2)
        
        agrupado = agrupado.sort_values(by='Gap', ascending=True)
        
        linha_total = {col: 'Total' for col in cols_group}
        linha_total.update({'N_valido': total_n, 'NPS': round(total_nps, 2), 'N_%_da_coluna': 100.0, 'Contribuição': round(total_nps, 2), 'Peso': 100.0, 'Gap': 0.0})
        return agrupado, linha_total

    if dim_principal is None:
        res, tot = processar_bloco(df, colunas_agrupamento)
        return pd.concat([res, pd.DataFrame([tot])], ignore_index=True) if not res.empty else pd.DataFrame()
    else:
        blocos = []
        for nome, subdf in df.groupby(dim_principal):
            res, tot = processar_bloco(subdf, colunas_agrupamento)
            if not res.empty:
                res.insert(0, dim_principal, nome)
                tot[dim_principal] = f"TOTAL - {nome}"
                blocos.append(pd.concat([res, pd.DataFrame([tot])], ignore_index=True))
        return pd.concat(blocos, ignore_index=True) if blocos else pd.DataFrame()

# 4. CARGA E TRATAMENTO ESPECÍFICO DE REVISÕES
print("Processando Ciclo de Revisões...")
df_pv = pd.read_excel(arquivo_pv, sheet_name='PV', skiprows=8).dropna(how='all')
if 'Unnamed: 0' in df_pv.columns: df_pv = df_pv.drop(columns=['Unnamed: 0'])

# Função para extrair o número da revisão e agrupar
def categorizar_revisao(texto):
    if pd.isna(texto): return None
    numeros = re.findall(r'\d+', str(texto))
    if not numeros: return None
    num = int(numeros[0])
    if num <= 4: return f"Revisão {num}"
    else: return "Revisão 5 ou +"

df_pv['Ciclo_Revisao'] = df_pv['Tipo de Entrevista'].apply(categorizar_revisao)
df_pv = df_pv[df_pv['Ciclo_Revisao'].notna()] # Filtra apenas o que é revisão

# Limpeza de strings
for col in ['Causa da nota de recomendação', 'Subcausa da nota de recomendação', 'Modelo']:
    df_pv[col] = df_pv[col].replace('-', 'Não especificada').fillna('Não especificada')

col_nps = 'Nota Recomendação concessionária (RESULTADO OFICIAL)'
df_pv[col_nps] = pd.to_numeric(df_pv[col_nps], errors='coerce')

# 5. EXECUÇÃO E CONSOLIDAÇÃO
tabelas_pbi = []
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    agr_causa = ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']
    
    # Processa por cada grupo de revisão
    for rev in sorted(df_pv['Ciclo_Revisao'].unique()):
        df_rev = df_pv[df_pv['Ciclo_Revisao'] == rev]
        nome_limpo = rev.replace(" ", "_")
        
        # 1. Causas daquela Revisão
        res_causa = gerar_analise_contribuicao(df_rev, agr_causa)
        res_causa.to_excel(writer, sheet_name=f"{nome_limpo}_Causa", index=False)
        tabelas_pbi.append(res_causa.assign(Aba_Origem=f"{nome_limpo}_Causa", Ciclo=rev))
        
        # 2. Modelos dentro daquela Revisão
        res_mod = gerar_analise_contribuicao(df_rev, agr_causa, dim_principal='Modelo')
        res_mod.to_excel(writer, sheet_name=f"{nome_limpo}_Modelo", index=False)
        tabelas_pbi.append(res_mod.assign(Aba_Origem=f"{nome_limpo}_Modelo", Ciclo=rev))

    # Aba mestre para o Power BI
    pd.concat(tabelas_pbi, ignore_index=True).to_excel(writer, sheet_name='Consolidado_PBI_Revisoes', index=False)

print(f"Relatório de Ciclo de Revisões gerado: {arquivo_saida}")