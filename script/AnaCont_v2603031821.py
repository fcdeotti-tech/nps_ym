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

# 1.1 Criar nome do arquivo de saída com Timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
arquivo_saida = os.path.join(caminho_pasta, f'Analise_NPS_Yamaha_{timestamp}.xlsx')

# 2. FUNÇÃO DE CÁLCULO DE NPS
def calcular_nps(notas):
    notas = pd.to_numeric(notas, errors='coerce').dropna()
    if len(notas) == 0:
        return np.nan
    promotores = len(notas[notas >= 9])
    detratores = len(notas[notas <= 6])
    nps = ((promotores - detratores) / len(notas)) * 100
    return nps

# 3. FUNÇÃO PARA CRIAR A TABELA ANALÍTICA (AGORA COM CÁLCULO RELATIVO E TOTAIS)
def gerar_analise_contribuicao(df, colunas_agrupamento, dim_principal=None, col_nps='Nota Recomendação concessionária (RESULTADO OFICIAL)'):
    """
    Se dim_principal (ex: 'Concessionária' ou 'Modelo') for fornecida, o cálculo será 
    relativo ao total DESSA dimensão, linha a linha, com subtotais por grupo.
    """
    
    # Função interna que processa o cálculo de um bloco (seja a base inteira, ou uma única loja)
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
        
        # Total Válido ignorando os "Não especificada"
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
        
        # Arredondamentos
        for col in ['N_%_da_coluna', 'N_%_Respostas_Validas', 'Peso', 'Contribuição', 'Gap', 'NPS']:
            agrupado[col] = agrupado[col].round(2)

        # Ordenar pelos piores Gaps primeiro (Ação prioritária)
        agrupado = agrupado.sort_values(by='Gap', ascending=True)
        
        # Construir a Linha de Total do Bloco
        linha_total = {col: 'Total' for col in cols_group}
        linha_total['N_valido'] = total_n
        linha_total['NPS'] = round(total_nps, 2)
        linha_total['N_%_da_coluna'] = 100.0
        linha_total['N_%_Respostas_Validas'] = 100.0 if total_respostas_validas_causa > 0 else np.nan
        linha_total['Contribuição'] = round(total_nps, 2)
        linha_total['Peso'] = 100.0 if total_nps != 0 else 0.0
        linha_total['Gap'] = 0.0
        
        return agrupado, linha_total

    # Se não tem dimensão principal (Análises Totais da VE/PV)
    if dim_principal is None:
        df_resultado, dict_total = processar_bloco(df, colunas_agrupamento)
        if df_resultado.empty: return pd.DataFrame()
        # Concatena a base com a linha de total embaixo
        return pd.concat([df_resultado, pd.DataFrame([dict_total])], ignore_index=True)
        
    # Se TEM dimensão principal (Por Concessionária ou Por Modelo)
    else:
        lista_blocos = []
        # Agrupa pelo dim_principal (que já ordena alfabeticamente por loja/modelo)
        for nome_grupo, df_grupo in df.groupby(dim_principal):
            df_resultado, dict_total = processar_bloco(df_grupo, colunas_agrupamento)
            if df_resultado.empty: continue
            
            # Insere a coluna da dimensão no início e assina a linha de Total
            df_resultado.insert(0, dim_principal, nome_grupo)
            dict_total[dim_principal] = f"TOTAL - {nome_grupo}"
            
            # Junta o bloco da loja com o total da loja
            bloco_completo = pd.concat([df_resultado, pd.DataFrame([dict_total])], ignore_index=True)
            lista_blocos.append(bloco_completo)
            
        if not lista_blocos: return pd.DataFrame()
        
        # Junta todos os blocos numa única tabela para o Excel
        return pd.concat(lista_blocos, ignore_index=True)

# 4. CARREGAR E LIMPAR DADOS
print("Carregando bases Excel...")
df_pv = pd.read_excel(arquivo_pv, sheet_name='PV', skiprows=8).dropna(how='all')
df_ve = pd.read_excel(arquivo_ve, sheet_name='VE', skiprows=8).dropna(how='all')

if 'Unnamed: 0' in df_pv.columns: df_pv = df_pv.drop(columns=['Unnamed: 0'])
if 'Unnamed: 0' in df_ve.columns: df_ve = df_ve.drop(columns=['Unnamed: 0'])

# Substituir "-" por "Não especificada"
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

# 5. GERAR ABA DE RESUMO DE DATAS SEPARADO
def gerar_resumo_datas(df, nome_base):
    colunas_data = [
        'Data do Evento', 'Data de Recebimento', 'Data de Envio', 
        'Data da Entrevista', 'Data de Atualização no portal'
    ]
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

df_resumo_ve = gerar_resumo_datas(df_ve, 'Vendas (VE)')
df_resumo_pv = gerar_resumo_datas(df_pv, 'Pós-Vendas (PV)')
df_resumo_datas = pd.concat([df_resumo_ve, df_resumo_pv], ignore_index=True)

# 6. REGRAS DE NEGÓCIO - TRATAMENTO ESPECÍFICO PÓS-VENDAS
def classificar_servico(tipo):
    if pd.isna(tipo): return 'Outros'
    tipo = str(tipo).upper()
    return 'Venda de Peças' if 'PÇS' in tipo or 'OUTROS' in tipo else 'Revisão'

df_pv['Categoria_PV'] = df_pv['Tipo de Entrevista'].apply(classificar_servico)

# 7. GERAR AS ANÁLISES NO EXCEL
print("Calculando contribuições e gerando o arquivo de saída...")
with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    
    # Salvar aba de Resumo
    df_resumo_datas.to_excel(writer, sheet_name='Resumo_Datas', index=False)
    
    agr_causa_sub = ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']
    agr_causa_only = ['Causa da nota de recomendação']
    
    # --- ANÁLISE DE VENDAS (VE) ---
    # Totais Globais
    gerar_analise_contribuicao(df_ve, agr_causa_only).to_excel(writer, sheet_name='VE_Total_Causa', index=False)
    gerar_analise_contribuicao(df_ve, agr_causa_sub).to_excel(writer, sheet_name='VE_Total_C_Sub', index=False)
    
    # Por Concessionária (Relativo à Concessionária)
    gerar_analise_contribuicao(df_ve, agr_causa_only, dim_principal='Concessionária').to_excel(writer, sheet_name='VE_Conc_Causa', index=False)
    gerar_analise_contribuicao(df_ve, agr_causa_sub, dim_principal='Concessionária').to_excel(writer, sheet_name='VE_Conc_C_Sub', index=False)

    # Por Modelo (Relativo ao Modelo)
    gerar_analise_contribuicao(df_ve, agr_causa_only, dim_principal='Modelo').to_excel(writer, sheet_name='VE_Mod_Causa', index=False)
    gerar_analise_contribuicao(df_ve, agr_causa_sub, dim_principal='Modelo').to_excel(writer, sheet_name='VE_Mod_C_Sub', index=False)

    # --- ANÁLISE DE PÓS-VENDAS (PV) ---
    for categoria in ['Revisão', 'Venda de Peças']:
        df_filtro_pv = df_pv[df_pv['Categoria_PV'] == categoria]
        cat_abrev = categoria[:5] 
        
        if not df_filtro_pv.empty:
            # Totais Globais da Categoria
            gerar_analise_contribuicao(df_filtro_pv, agr_causa_only).to_excel(writer, sheet_name=f'PV_{cat_abrev}_Tot_Causa', index=False)
            gerar_analise_contribuicao(df_filtro_pv, agr_causa_sub).to_excel(writer, sheet_name=f'PV_{cat_abrev}_Tot_C_Sub', index=False)
            
            # Por Concessionária (Relativo à Concessionária dentro da Categoria)
            gerar_analise_contribuicao(df_filtro_pv, agr_causa_only, dim_principal='Concessionária').to_excel(writer, sheet_name=f'PV_{cat_abrev}_Conc_Causa', index=False)
            gerar_analise_contribuicao(df_filtro_pv, agr_causa_sub, dim_principal='Concessionária').to_excel(writer, sheet_name=f'PV_{cat_abrev}_Conc_C_Sub', index=False)

print(f"\nRelatório gerado com sucesso! Arquivo salvo em:\n{arquivo_saida}")