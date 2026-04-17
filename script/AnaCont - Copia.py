import pandas as pd
import numpy as np
import warnings
import os
from datetime import datetime

warnings.filterwarnings('ignore')

# 1. DEFINIÇÃO DOS CAMINHOS LOCAIS
caminho_pasta = r"C:\Users\fernando.deotti\OneDrive - Holding Specialist Researchers Brasil LTDA\Documentos\Projetos\Yamaha\NPS\Python"
arquivo_pv = os.path.join(caminho_pasta, 'PV.xlsx')
arquivo_ve = os.path.join(caminho_pasta, 'VE.xlsx')

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
arquivo_saida = os.path.join(caminho_pasta, f'Analise_NPS_Yamaha_{timestamp}.xlsx')

# 2. FUNÇÃO DE CÁLCULO DE NPS
def calcular_nps(notas):
    notas = pd.to_numeric(notas, errors='coerce').dropna()
    if len(notas) == 0:
        return np.nan
    promotores = len(notas[notas >= 9])
    detratores = len(notas[notas <= 6])
    return ((promotores - detratores) / len(notas)) * 100

# 3. FUNÇÃO PARA CRIAR A TABELA ANALÍTICA
def gerar_analise_contribuicao(df, colunas_agrupamento, dim_principal=None, col_nps='Nota Recomendação concessionária (RESULTADO OFICIAL)'):
    def processar_bloco(df_bloco, cols_group):
        total_n = len(df_bloco[col_nps].dropna())
        total_nps = calcular_nps(df_bloco[col_nps])
        
        if total_n == 0:
            return pd.DataFrame(), None
            
        agrupado = df_bloco.groupby(cols_group).agg(
            N_valido=(col_nps, lambda x: x.notna().sum()),
            NPS=(col_nps, calcular_nps)
        ).reset_index()

        agrupado = agrupado[agrupado['N_valido'] > 0]
        
        mascara_valida = pd.Series(True, index=agrupado.index)
        for col in cols_group:
            if col in agrupado.columns:
                mascara_valida = mascara_valida & (agrupado[col] != 'Não especificada') & (agrupado[col].notna()) & (agrupado[col] != '')
        
        total_respostas_validas_causa = agrupado.loc[mascara_valida, 'N_valido'].sum()
        
        agrupado['N_%_da_coluna'] = (agrupado['N_valido'] / total_n) * 100
        
        def calc_respostas_validas(row):
            is_valid = True
            for col in cols_group:
                if row.get(col) == 'Não especificada' or pd.isna(row.get(col)) or row.get(col) == '':
                    is_valid = False
            if is_valid and total_respostas_validas_causa > 0:
                return (row['N_valido'] / total_respostas_validas_causa) * 100
            else:
                return np.nan
                
        agrupado['N_%_Respostas_Validas'] = agrupado.apply(calc_respostas_validas, axis=1)
        
        contribuicao_real = (agrupado['N_valido'] / total_n) * agrupado['NPS']
        agrupado['Contribuição'] = contribuicao_real
        agrupado['Peso'] = (contribuicao_real / total_nps) * 100 if total_nps != 0 else 0
        agrupado['Gap'] = agrupado['Peso'] - agrupado['N_%_da_coluna']
        
        for col in ['N_%_da_coluna', 'N_%_Respostas_Validas', 'Peso', 'Contribuição', 'Gap', 'NPS']:
            agrupado[col] = agrupado[col].round(2)

        agrupado = agrupado.sort_values(by='Gap', ascending=True)
        
        linha_total = {col: 'Total' for col in cols_group}
        linha_total['N_valido'] = total_n
        linha_total['NPS'] = round(total_nps, 2)
        linha_total['N_%_da_coluna'] = 100.0
        linha_total['N_%_Respostas_Validas'] = 100.0 if total_respostas_validas_causa > 0 else np.nan
        linha_total['Contribuição'] = round(total_nps, 2)
        linha_total['Peso'] = 100.0 if total_nps != 0 else 0.0
        linha_total['Gap'] = 0.0
        
        return agrupado, linha_total

    if dim_principal is None:
        df_resultado, dict_total = processar_bloco(df, colunas_agrupamento)
        if df_resultado.empty: return pd.DataFrame()
        return pd.concat([df_resultado, pd.DataFrame([dict_total])], ignore_index=True)
    else:
        lista_blocos = []
        for nome_grupo, df_grupo in df.groupby(dim_principal):
            df_resultado, dict_total = processar_bloco(df_grupo, colunas_agrupamento)
            if df_resultado.empty: continue
            
            df_resultado.insert(0, dim_principal, nome_grupo)
            dict_total[dim_principal] = f"TOTAL - {nome_grupo}"
            
            bloco_completo = pd.concat([df_resultado, pd.DataFrame([dict_total])], ignore_index=True)
            lista_blocos.append(bloco_completo)
            
        if not lista_blocos: return pd.DataFrame()
        return pd.concat(lista_blocos, ignore_index=True)

# 4. CARREGAR E LIMPAR DADOS
print("Carregando bases Excel...")
df_pv = pd.read_excel(arquivo_pv, sheet_name='PV', skiprows=8).dropna(how='all')
df_ve = pd.read_excel(arquivo_ve, sheet_name='VE', skiprows=8).dropna(how='all')

if 'Unnamed: 0' in df_pv.columns: df_pv = df_pv.drop(columns=['Unnamed: 0'])
if 'Unnamed: 0' in df_ve.columns: df_ve = df_ve.drop(columns=['Unnamed: 0'])

colunas_ajuste_texto = ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']
for col in colunas_ajuste_texto:
    if col in df_pv.columns:
        df_pv[col] = df_pv[col].replace('-', 'Não especificada').fillna('Não especificada')
        df_pv[col] = df_pv[col].apply(lambda x: 'Não especificada' if str(x).strip() == '' else x)
    if col in df_ve.columns:
        df_ve[col] = df_ve[col].replace('-', 'Não especificada').fillna('Não especificada')
        df_ve[col] = df_ve[col].apply(lambda x: 'Não especificada' if str(x).strip() == '' else x)

col_nps = 'Nota Recomendação concessionária (RESULTADO OFICIAL)'
df_pv[col_nps] = pd.to_numeric(df_pv[col_nps], errors='coerce')
df_ve[col_nps] = pd.to_numeric(df_ve[col_nps], errors='coerce')

# 5. GERAR ABA DE RESUMO DE DATAS
def gerar_resumo_datas(df, nome_base):
    colunas_data = ['Data do Evento', 'Data de Recebimento', 'Data de Envio', 'Data da Entrevista', 'Data de Atualização no portal']
    resumo = []
    for col in colunas_data:
        if col in df.columns:
            datas_convertidas = pd.to_datetime(df[col], format='%d/%m/%Y', errors='coerce')
            min_data = datas_convertidas.min()
            max_data = datas_convertidas.max()
            resumo.append({
                'Base': nome_base,
                'Campo de Data': col,
                'Data Mínima': min_data.strftime('%d/%m/%Y') if pd.notna(min_data) else 'N/A',
                'Data Máxima': max_data.strftime('%d/%m/%Y') if pd.notna(max_data) else 'N/A'
            })
    return pd.DataFrame(resumo)

df_resumo_datas = pd.concat([gerar_resumo_datas(df_ve, 'Vendas (VE)'), gerar_resumo_datas(df_pv, 'Pós-Vendas (PV)')], ignore_index=True)

# 6. REGRAS DE NEGÓCIO - PÓS-VENDAS
def classificar_servico(tipo):
    if pd.isna(tipo): return 'Outros'
    tipo = str(tipo).upper()
    return 'Venda de Peças' if 'PÇS' in tipo or 'OUTROS' in tipo else 'Revisão'

df_pv['Categoria_PV'] = df_pv['Tipo de Entrevista'].apply(classificar_servico)

# 7. GERAR AS ANÁLISES E CONSOLIDAR PARA O POWER BI
print("Calculando contribuições e gerando o arquivo de saída...")

# Lista que vai guardar todos os dataframes para empilhar no final
tabelas_para_pbi = []

# Função auxiliar para salvar na aba individual E adicionar na lista master
def processar_e_salvar(df_resultado, nome_aba, writer_obj):
    if not df_resultado.empty:
        # 1. Salva a aba individual (para consulta rápida em Excel)
        df_resultado.to_excel(writer_obj, sheet_name=nome_aba, index=False)
        
        # 2. Prepara para o Power BI
        df_pbi = df_resultado.copy()
        # Insere a coluna de origem como a primeira coluna
        df_pbi.insert(0, 'Aba_Origem', nome_aba)
        tabelas_para_pbi.append(df_pbi)

with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    
    df_resumo_datas.to_excel(writer, sheet_name='Resumo_Datas', index=False)
    
    agr_causa_sub = ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']
    agr_causa_only = ['Causa da nota de recomendação']
    
    # --- VENDAS (VE) ---
    processar_e_salvar(gerar_analise_contribuicao(df_ve, agr_causa_only), 'VE_Total_Causa', writer)
    processar_e_salvar(gerar_analise_contribuicao(df_ve, agr_causa_sub), 'VE_Total_C_Sub', writer)
    processar_e_salvar(gerar_analise_contribuicao(df_ve, agr_causa_only, 'Concessionária'), 'VE_Conc_Causa', writer)
    processar_e_salvar(gerar_analise_contribuicao(df_ve, agr_causa_sub, 'Concessionária'), 'VE_Conc_C_Sub', writer)
    processar_e_salvar(gerar_analise_contribuicao(df_ve, agr_causa_only, 'Modelo'), 'VE_Mod_Causa', writer)
    processar_e_salvar(gerar_analise_contribuicao(df_ve, agr_causa_sub, 'Modelo'), 'VE_Mod_C_Sub', writer)

    # --- PÓS-VENDAS (PV) ---
    for categoria in ['Revisão', 'Venda de Peças']:
        df_filtro_pv = df_pv[df_pv['Categoria_PV'] == categoria]
        cat_abrev = categoria[:5] 
        
        processar_e_salvar(gerar_analise_contribuicao(df_filtro_pv, agr_causa_only), f'PV_{cat_abrev}_Tot_Causa', writer)
        processar_e_salvar(gerar_analise_contribuicao(df_filtro_pv, agr_causa_sub), f'PV_{cat_abrev}_Tot_C_Sub', writer)
        processar_e_salvar(gerar_analise_contribuicao(df_filtro_pv, agr_causa_only, 'Concessionária'), f'PV_{cat_abrev}_Conc_Causa', writer)
        processar_e_salvar(gerar_analise_contribuicao(df_filtro_pv, agr_causa_sub, 'Concessionária'), f'PV_{cat_abrev}_Conc_C_Sub', writer)

    # --- GERAR ABA CONSOLIDADA PARA O POWER BI ---
    if tabelas_para_pbi:
        df_consolidado = pd.concat(tabelas_para_pbi, ignore_index=True)
        # Salva a base consolidada como a segunda aba do arquivo
        df_consolidado.to_excel(writer, sheet_name='Base_Consolidada_PBI', index=False)

print(f"\nRelatório gerado com sucesso! Arquivo salvo em:\n{arquivo_saida}")