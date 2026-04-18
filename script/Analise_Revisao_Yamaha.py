import pandas as pd
import numpy as np
import warnings
import os
import re
from datetime import datetime

warnings.filterwarnings('ignore')

# ==========================================
# 1. DEFINIÇÃO DOS CAMINHOS (WSL / LINUX)
# ==========================================
DIRETORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
DIRETORIO_RAIZ = os.path.dirname(DIRETORIO_SCRIPT)

DIRETORIO_SOURCE = os.path.join(DIRETORIO_RAIZ, 'source')
DIRETORIO_OUTPUT = os.path.join(DIRETORIO_RAIZ, 'output')

# Nota: Revisão só usa a base de Pós-Vendas (PV)
arquivo_pv = os.path.join(DIRETORIO_SOURCE, 'PV.xlsx')

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
arquivo_saida = os.path.join(DIRETORIO_OUTPUT, f'Analise_Revisoes_Yamaha_{timestamp}.xlsx')

# ==========================================
# 2. FUNÇÃO DE CÁLCULO DE NPS
# ==========================================
def calcular_nps(notas):
    notas = pd.to_numeric(notas, errors='coerce').dropna()
    if len(notas) == 0: return np.nan
    promotores = len(notas[notas >= 9])
    detratores = len(notas[notas <= 6])
    return ((promotores - detratores) / len(notas)) * 100

# ==========================================
# 3. FUNÇÃO ANALÍTICA MESTRE (Com injeção de TOTAL DA CAUSA)
# ==========================================
def gerar_analise_contribuicao(df, colunas_agrupamento, dim_principal=None, col_nps='Nota Recomendação concessionária (RESULTADO OFICIAL)'):
    def processar_bloco(df_bloco, cols_group):
        total_n = len(df_bloco[col_nps].dropna())
        total_nps = calcular_nps(df_bloco[col_nps])
        if total_n == 0: return pd.DataFrame(), None
            
        def agrupar_e_calcular(cols):
            grp = df_bloco.groupby(cols).agg(
                N_valido=(col_nps, lambda x: x.notna().sum()),
                NPS=(col_nps, calcular_nps)
            ).reset_index()
            grp = grp[grp['N_valido'] > 0]
            grp['N_%_da_coluna'] = (grp['N_valido'] / total_n) * 100
            contrib = (grp['N_valido'] / total_n) * grp['NPS']
            grp['Contribuição'] = contrib
            grp['Peso'] = (contrib / total_nps) * 100 if total_nps != 0 else 0
            grp['Gap'] = grp['Peso'] - grp['N_%_da_coluna']
            for c in ['N_%_da_coluna', 'Peso', 'Contribuição', 'Gap', 'NPS']:
                grp[c] = grp[c].round(2)
            return grp

        agrupado = agrupar_e_calcular(cols_group)
        
        # INJEÇÃO DO TOTAL DA CAUSA E ORDENAÇÃO INTELIGENTE
        if 'Causa da nota de recomendação' in cols_group and 'Subcausa da nota de recomendação' in cols_group:
            cols_sem_sub = [c for c in cols_group if c != 'Subcausa da nota de recomendação']
            
            # Calcula os totais apenas para as Causas
            agr_causa = agrupar_e_calcular(cols_sem_sub)
            agr_causa['Subcausa da nota de recomendação'] = 'TOTAL DA CAUSA'
            
            # Prepara um mapa com o Gap Total de cada causa para forçar a ordenação em blocos
            mapa_gap_causa = agr_causa[cols_sem_sub + ['Gap']].rename(columns={'Gap': 'Gap_Total_Causa'})
            
            # Junta os dados normais com as linhas de TOTAL DA CAUSA
            agrupado = pd.concat([agrupado, agr_causa], ignore_index=True)
            
            # Traz a coluna Gap_Total_Causa para todas as linhas
            agrupado = agrupado.merge(mapa_gap_causa, on=cols_sem_sub, how='left')
            
            # Identifica qual linha é a totalizadora
            agrupado['is_total_causa'] = agrupado['Subcausa da nota de recomendação'] == 'TOTAL DA CAUSA'
            
            # Ordena por: Pior Gap de Causa, depois a linha TOTAL DA CAUSA no topo, e por fim o pior Gap de Subcausa
            agrupado = agrupado.sort_values(
                by=['Gap_Total_Causa', 'Causa da nota de recomendação', 'is_total_causa', 'Gap'], 
                ascending=[True, True, False, True]
            )
            
            # Limpa colunas auxiliares de ordenação
            agrupado = agrupado.drop(columns=['Gap_Total_Causa', 'is_total_causa'])
            
        else:
            # Se for uma tabela sem subcausas, ordena normalmente pelo Gap
            agrupado = agrupado.sort_values(by='Gap', ascending=True)

        # LINHA DE TOTAL FINAL DA VISÃO
        linha_total = {col: 'Total' for col in cols_group}
        linha_total.update({'N_valido': total_n, 'NPS': round(total_nps, 2), 'N_%_da_coluna': 100.0, 'Contribuição': round(total_nps, 2), 'Peso': 100.0, 'Gap': 0.0})
        
        return agrupado, linha_total

    # Processamento e injeção do Dim Principal (Ex: Modelo)
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

# ==========================================
# 4. CARGA E TRATAMENTO ESPECÍFICO DE REVISÕES
# ==========================================
print("Carregando base de PV e processando Ciclo de Revisões...")
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

# ==========================================
# 5. EXECUÇÃO E CONSOLIDAÇÃO NO EXCEL
# ==========================================
print("Calculando Gaps com injeção de Totais e exportando abas...")
tabelas_pbi = []

with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    agr_causa = ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']
    
    # Processa por cada grupo de revisão (1, 2, 3, 4, 5 ou +)
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

    # Aba mestre consolidada para o Power BI
    pd.concat(tabelas_pbi, ignore_index=True).to_excel(writer, sheet_name='Consolidado_PBI_Revisoes', index=False)

print(f"\nRelatório de Ciclo de Revisões concluído com sucesso!\nArquivo guardado em: {DIRETORIO_OUTPUT}")