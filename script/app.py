import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import re

# ==========================================
# IDENTIDADE VISUAL INSIGHTS&ETC (CSS)
# ==========================================
st.set_page_config(page_title="Yamaha NPS Explorer", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] { font-family: 'Montserrat', sans-serif; }
    
    /* Removemos a cor fixa daqui para respeitar o Dark Mode nativo do usuário */
    h1, h2, h3 { font-family: 'Montserrat', sans-serif !important; font-weight: 700 !important; }
    
    section[data-testid="stSidebar"] { background-color: #0A192F; color: white; }
    section[data-testid="stSidebar"] .stMarkdown p { color: white; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# FUNÇÃO DE CORREÇÃO DO GRÁFICO DE IMPACTO
# ==========================================
# ==========================================
# FUNÇÃO DE CORREÇÃO DO GRÁFICO DE IMPACTO
# ==========================================
def gerar_grafico_impacto_corrigido(df_plot, col_y):
    """
    Usa graph_objects para controle absoluto.
    Força eixo simétrico, escala RdYlGn nativa e impede corte de texto negativo.
    """
    df_plot = df_plot.copy()
    
    # 1. Garantir conversão numérica forte e ordenação
    df_plot['Impacto'] = pd.to_numeric(df_plot['Impacto'], errors='coerce').fillna(0.0)
    df_plot = df_plot.sort_values('Impacto', ascending=True)
    
    # 2. Definir o limite simétrico do eixo X (baseado no maior valor absoluto)
    max_abs = df_plot['Impacto'].abs().max()
    max_val = (max_abs * 1.3) if pd.notna(max_abs) and max_abs > 0 else 1.0

    # 3. Construção via Graph Objects (Extremamente robusto no GCP)
    fig = go.Figure()
    
    fig.add_trace(go.Bar(
        x=df_plot['Impacto'],
        y=df_plot[col_y],
        orientation='h',
        text=[f"{v:.1f}" for v in df_plot['Impacto']], # Formata o texto manualmente
        textposition='outside',
        cliponaxis=False,
        marker=dict(
            color=df_plot['Impacto'],
            colorscale='RdYlGn',
            cmid=0, # Garante que o zero é a transição de cor
            line=dict(width=0)
        )
    ))
    
    fig.update_layout(
        yaxis=dict(
            categoryorder='array', 
            categoryarray=df_plot[col_y].tolist() # Mantém a ordem certa das barras
        ),
        xaxis=dict(
            range=[-max_val, max_val], # Força a linha do zero a ficar no meio
            zeroline=True,
            zerolinewidth=2,
            zerolinecolor='rgba(0,0,0,0.5)',
            showgrid=False # Design limpo sem linhas de fundo
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        # O SEGREDO AQUI: 'l=50' (Margem esquerda) impede que os negativos sumam da tela!
        margin=dict(r=50, l=50, t=30, b=30) 
    )
    
    return fig

# ==========================================
# 1. SISTEMA DE LOGIN NATIVO
# ==========================================
USUARIOS = {
    "diretoria_yamaha": {"nome": "Diretoria Yamaha", "senha": "yamaha_nps_2026"},
    "gerencia_vendas": {"nome": "Gerência de Vendas", "senha": "vendas_yamaha"},
    "gerencia_pos_vendas": {"nome": "Gerência de Pós-Vendas", "senha": "posvendas_yamaha"},
    "admin": {"nome": "Fernando Deotti", "senha": "root_specialist"}
}

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

def verificar_login():
    usuario = st.session_state["campo_usuario"]
    senha = st.session_state["campo_senha"]
    
    if usuario in USUARIOS and USUARIOS[usuario]["senha"] == senha:
        st.session_state["autenticado"] = True
        st.session_state["nome_usuario"] = USUARIOS[usuario]["nome"]
        st.session_state["erro_login"] = False
    else:
        st.session_state["erro_login"] = True

def fazer_logout():
    st.session_state["autenticado"] = False
    st.session_state["nome_usuario"] = ""

if not st.session_state["autenticado"]:
    col1, col2 = st.columns([1, 15])
    with col1:
        st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Yamaha_Motor_Logo.svg/1024px-Yamaha_Motor_Logo.svg.png", width=70)
    with col2:
        st.title("Yamaha NPS Explorer")
    
    st.markdown("---")
    col_login, col_vazia = st.columns([1, 2])
    with col_login:
        st.subheader("Acesso Restrito")
        st.text_input("Usuário", key="campo_usuario")
        st.text_input("Senha", type="password", key="campo_senha")
        st.button("Entrar", on_click=verificar_login)
        
        if st.session_state.get("erro_login"):
            st.error("Usuário ou senha incorretos. Tente novamente.")
            
    st.stop()

# ==========================================
# 2. EXTRATOR NUMÉRICO (A Britadeira)
# ==========================================
def extrair_numero(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    
    val_str = str(val).replace(',', '.').replace('−', '-')
    val_str = re.sub(r'[^\d\.\-]', '', val_str)
    
    try:
        return float(val_str) if val_str and val_str != '-' else 0.0
    except ValueError:
        return 0.0

# ==========================================
# 3. MAPEAMENTO E LEITURA 
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(os.path.dirname(BASE_DIR), 'output')

@st.cache_data
def ler_dados_nps_oficial(file_name, sheet_name):
    path = os.path.join(OUTPUT_DIR, file_name)
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
            df.columns = df.columns.str.strip() 
            
            colunas_alvo = [
                'NPS', 'Gap', 'Impacto', 'Contribuição', 'Peso', 'N_valido', 'Respondentes', 
                'Volume (N)', 'Volume', 'NPS Atual', 'NPS Potencial', 'Ganho Possível'
            ]
            
            for col in colunas_alvo:
                if col in df.columns:
                    df[col] = df[col].apply(extrair_numero)
                    
            return df
        except Exception as e:
            st.error(f"Erro ao ler a aba '{sheet_name}': {e}")
            return pd.DataFrame()
    return pd.DataFrame()

def get_top_bottom_10(df_agrupado, col_valor):
    df_sorted = df_agrupado.sort_values(col_valor, ascending=True)
    if len(df_sorted) > 20:
        df_bottom = df_sorted.head(10)
        df_top = df_sorted.tail(10)
        return pd.concat([df_bottom, df_top])
    return df_sorted

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

def exibir_tamanho_amostra(df):
    if df.empty: return
    df_calc = df.copy()
    
    if 'Causa da nota de recomendação' in df_calc.columns: df_calc = df_calc[~df_calc['Causa da nota de recomendação'].astype(str).str.upper().isin(['TOTAL'])]
    if 'Subcausa da nota de recomendação' in df_calc.columns: df_calc = df_calc[~df_calc['Subcausa da nota de recomendação'].astype(str).str.upper().isin(['TOTAL DA CAUSA', 'TOTAL'])]
    if 'Região' in df_calc.columns: df_calc = df_calc[~df_calc['Região'].astype(str).str.upper().isin(['TOTAL NACIONAL', 'TOTAL'])]
    if 'Grupo' in df_calc.columns: df_calc = df_calc[~df_calc['Grupo'].astype(str).str.upper().str.startswith('TOTAL -')]
    if 'Concessionária' in df_calc.columns: df_calc = df_calc[~df_calc['Concessionária'].astype(str).str.upper().str.startswith('TOTAL -')]
    if 'Modelo' in df_calc.columns: df_calc = df_calc[~df_calc['Modelo'].astype(str).str.upper().str.startswith('TOTAL -')]
        
    n = 0
    if 'N_valido' in df_calc.columns: n = df_calc['N_valido'].sum()
    elif 'Respondentes' in df_calc.columns: n = df_calc['Respondentes'].sum()
    elif 'Volume (N)' in df_calc.columns: n = df_calc['Volume (N)'].sum()
    elif 'Volume' in df_calc.columns: n = df_calc['Volume'].sum()
        
    st.markdown(f"<div style='text-align: right; font-size: 13px; color: #888; margin-bottom: 5px;'>Tamanho da Amostra (N): <b>{int(n):,}</b></div>".replace(',', '.'), unsafe_allow_html=True)

def mostrar_tabela_formatada(df, filename_download="dados.xlsx"):
    if df.empty:
        st.info("Nenhum dado encontrado para os filtros selecionados.")
        return

    df_display = df.copy()
    
    if 'Gap' in df_display.columns and 'Impacto' in df_display.columns:
        df_display = df_display.drop(columns=['Impacto'])
        
    rename_dict = {'N_valido': 'Respondentes', 'N_%_da_coluna': '% Respondentes', 'N_%_Respostas_Validas': '% Respostas válidas', 'Gap': 'Impacto'}
    df_display = df_display.rename(columns=rename_dict)
    
    termos_limpar = ['-', 'Não especificada', 'Total', 'TOTAL DA CAUSA']
    for col in ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']:
        if col in df_display.columns:
            df_display[col] = df_display[col].replace('Nenhuma das opções acima', 'Nenhuma subcausa específica')
            df_display = df_display[~df_display[col].isin(termos_limpar)]

    cols = list(df_display.columns)
    if 'Respondentes' in cols and '% Respondentes' in cols:
        cols.remove('% Respondentes')
        cols.insert(cols.index('Respondentes') + 1, '% Respondentes')
        df_display = df_display[cols]

    col_num = ['NPS', 'Impacto', 'Contribuição', 'Peso', '% Respondentes', '% Respostas válidas', 'NPS Atual', 'NPS Potencial', 'Ganho Possível']
    fmt = {c: lambda x: f"{x:.1f}" if pd.notna(x) else "-" for c in col_num if c in df_display.columns}
    
    st.dataframe(df_display.style.format(fmt), use_container_width=True, hide_index=True)
    st.download_button("📥 Baixar Excel", convert_df_to_excel(df_display), filename_download, "application/vnd.ms-excel")

# ==========================================
# 4. SIDEBAR - FILTROS GLOBAIS
# ==========================================
st.sidebar.markdown("<h1>🏍️</h1>", unsafe_allow_html=True)
st.sidebar.markdown(f"**Olá, {st.session_state['nome_usuario']}!**")
if st.sidebar.button("Sair (Logout)"):
    fazer_logout()
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()
        
st.sidebar.markdown("---")
departamento = st.sidebar.radio("Dep", ["Vendas (VE)", "Pós-Vendas (PV)"], label_visibility="collapsed")
dep_prefix = "VE" if departamento == "Vendas (VE)" else "PV"

tipo_analise = st.sidebar.selectbox("Tipo de Análise", ["Contribuição Total", "Análise de Neutros", "Análise de Detratores", "Ciclo de Revisões"])

st.sidebar.markdown("---")
st.sidebar.subheader("Filtros Globais")

df_filtros = ler_dados_nps_oficial("Analise_NPS_Yamaha.xlsx", f"{dep_prefix}_Tot_C_Sub")

def limpar(lista):
    return sorted([str(x) for x in lista if pd.notna(x) and str(x) not in ['-', 'Total', 'TOTAL DA CAUSA', 'Não especificada']])

if not df_filtros.empty:
    f_causa = st.sidebar.multiselect("Causa", limpar(df_filtros['Causa da nota de recomendação'].unique()))
    f_sub = st.sidebar.multiselect("Subcausa", limpar(df_filtros['Subcausa da nota de recomendação'].unique()))
else:
    f_causa, f_sub = [], []

f_nps = st.sidebar.multiselect("Faixa de NPS", ["Menor que 0", "0 a 49", "50 a 74", "75 a 100"])
f_imp = st.sidebar.multiselect("Impacto", ["Positivo", "Negativo"])

def aplicar_filtros_globais(df):
    if df.empty: return df
    
    if 'Região' in df.columns: 
        df = df[~df['Região'].astype(str).str.upper().str.contains('PERFIL', na=False)]
    if f_causa and 'Causa da nota de recomendação' in df.columns: 
        df = df[df['Causa da nota de recomendação'].isin(f_causa)]
    if f_sub and 'Subcausa da nota de recomendação' in df.columns: 
        df = df[df['Subcausa da nota de recomendação'].isin(f_sub)]
    
    if f_nps:
        col = 'NPS' if 'NPS' in df.columns else ('NPS Atual' if 'NPS Atual' in df.columns else None)
        if col:
            val = pd.to_numeric(df[col], errors='coerce')
            m = []
            if "Menor que 0" in f_nps: m.append(val < 0)
            if "0 a 49" in f_nps: m.append((val >= 0) & (val < 50))
            if "50 a 74" in f_nps: m.append((val >= 50) & (val < 75))
            if "75 a 100" in f_nps: m.append(val >= 75)
            if m:
                mask = m[0]
                for i in m[1:]: mask |= i
                df = df[mask]

    if f_imp:
        col = 'Gap' if 'Gap' in df.columns else ('Impacto' if 'Impacto' in df.columns else ('Ganho Possível' if 'Ganho Possível' in df.columns else None))
        if col:
            val = pd.to_numeric(df[col], errors='coerce')
            if "Positivo" in f_imp and "Negativo" not in f_imp: df = df[val >= 0]
            elif "Negativo" in f_imp and "Positivo" not in f_imp: df = df[val < 0]
            
    return df

termos_omitir_graficos = ['-', 'Não especificada', 'Nenhuma das opções acima', 'Nenhuma subcausa específica', 'Total', 'TOTAL DA CAUSA']

# ==========================================
# 5. LÓGICA DE VISUALIZAÇÃO POR TIPO
# ==========================================

if tipo_analise == "Contribuição Total":
    st.title(f"📊 Contribuição e Impactos: {departamento}")
    t1, t2, t3, t4, t5 = st.tabs(["Nacional", "Regiões", "Grupos", "Concessionárias", "Modelos"])
    file = "Analise_NPS_Yamaha.xlsx"

    with t1:
        df_c = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Tot_Causa"))
        df_s = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Tot_C_Sub"))
        
        exibir_tamanho_amostra(df_s) 
        
        if not df_c.empty:
            st.subheader("Ponte de NPS: Contribuição por Causa")
            df_wf = df_c[~df_c['Causa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')].copy()
            df_wf['Causa da nota de recomendação'] = df_wf['Causa da nota de recomendação'].replace(['-', 'Não especificada'], 'Não indicou motivo específico')
            
            ordem_wf = ["Não indicou motivo específico", "Instalações físicas", "Equipe de consultores" if dep_prefix == "PV" else "Equipe de vendedores", "Momento da negociação", "Entrega técnica", "Outro(s) motivo(s)"]
            df_wf['Causa da nota de recomendação'] = pd.Categorical(df_wf['Causa da nota de recomendação'], categories=ordem_wf, ordered=True)
            df_wf = df_wf.sort_values('Causa da nota de recomendação').dropna(subset=['Causa da nota de recomendação'])
            
            if not df_wf.empty:
                x_vals = list(df_wf['Causa da nota de recomendação']) + ["NPS Total"]
                deltas = list(df_wf['Contribuição'])
                y_vals = deltas + [sum(deltas)]
                
                bases, atual = [], 0
                for d in deltas:
                    bases.append(atual)
                    atual += d
                bases.append(0)

                colors = []
                for cat, val in zip(x_vals, y_vals):
                    if cat == "NPS Total": colors.append("#1f77b4")
                    elif cat in ["Não indicou motivo específico", "Outro(s) motivo(s)"]: colors.append("#808080")
                    elif val >= 0: colors.append("#00D2D3") # CorInsights Positivo
                    else: colors.append("#FF6B6B") # CorInsights Negativo

                texts = [f"{v:.1f}" for v in y_vals]
                fig_wf = go.Figure()
                fig_wf.add_trace(go.Bar(x=x_vals, y=y_vals, base=bases, marker_color=colors, text=texts, textposition='outside'))
                for i in range(len(x_vals) - 1):
                    fig_wf.add_shape(type="line", x0=i - 0.4, y0=bases[i+1], x1=i + 1.4, y1=bases[i+1], line=dict(color="rgba(63, 63, 63, 0.4)", width=1, dash="dot"))

                fig_wf.update_layout(title="Como cada causa constrói o NPS Final", showlegend=False)
                st.plotly_chart(fig_wf, use_container_width=True)

            if not df_s.empty:
                st.subheader("Matriz de Impacto por Subcausa")
                df_gap = df_s[~df_s['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)].copy()
                
                if 'Gap' in df_gap.columns:
                    if 'Impacto' in df_gap.columns: df_gap = df_gap.drop(columns=['Impacto'])
                    df_gap = df_gap.rename(columns={'Gap': 'Impacto'})
                    
                    if not df_gap.empty:
                        fig_gap = gerar_grafico_impacto_corrigido(df_gap, 'Subcausa da nota de recomendação')
                        st.plotly_chart(fig_gap, use_container_width=True)

            st.subheader("Impacto Consolidado por Causa")
            mostrar_tabela_formatada(df_c, "Nacional_Causa.xlsx")
            st.subheader("Detalhe por Subcausa")
            mostrar_tabela_formatada(df_s, "Nacional_Subcausa.xlsx")

    with t2:
        df = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Reg_C_Sub" if dep_prefix=="VE" else f"PV_Tot_Reg_C_Sub"))
        if not df.empty:
            sel = st.multiselect("Filtrar Região", limpar(df['Região'].unique()), key="reg_t2")
            if sel: df = df[df['Região'].isin(sel)]
            
            exibir_tamanho_amostra(df)
            df_chart = df[~df['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)].copy()
            if not df_chart.empty and 'Gap' in df_chart.columns:
                st.subheader("Visão Geral de Impacto")
                df_chart['Gap'] = pd.to_numeric(df_chart['Gap'], errors='coerce').fillna(0.0)
                df_chart_agg = df_chart.groupby('Região')['Gap'].sum().reset_index()
                df_chart_agg = df_chart_agg.rename(columns={'Gap': 'Impacto'})
                
                fig = gerar_grafico_impacto_corrigido(get_top_bottom_10(df_chart_agg, 'Impacto'), 'Região')
                st.plotly_chart(fig, use_container_width=True)
                
            mostrar_tabela_formatada(df.sort_values(['Região', 'Gap' if 'Gap' in df.columns else 'Região'], ascending=[True, False]), "Regioes.xlsx")

    with t3:
        df = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Grup_C_Sub" if dep_prefix=="VE" else f"PV_Tot_Grup_C_Sub"))
        if not df.empty:
            c1, c2 = st.columns(2)
            with c1: r_sel = st.multiselect("Filtrar por Região", limpar(df['Região'].unique() if 'Região' in df.columns else []), key="reg_g")
            with c2: g_sel = st.multiselect("Filtrar por Grupo", limpar(df['Grupo'].unique()), key="grp_g")
            if r_sel and 'Região' in df.columns: df = df[df['Região'].isin(r_sel)]
            if g_sel: df = df[df['Grupo'].isin(g_sel)]
            
            exibir_tamanho_amostra(df)
            df_chart = df[~df['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)].copy()
            if not df_chart.empty and 'Gap' in df_chart.columns:
                st.subheader("Top 10 / Bottom 10 Impactos por Grupo")
                df_chart['Gap'] = pd.to_numeric(df_chart['Gap'], errors='coerce').fillna(0.0)
                df_chart_agg = df_chart.groupby('Grupo')['Gap'].sum().reset_index()
                df_chart_agg = df_chart_agg.rename(columns={'Gap': 'Impacto'})
                
                fig = gerar_grafico_impacto_corrigido(get_top_bottom_10(df_chart_agg, 'Impacto'), 'Grupo')
                st.plotly_chart(fig, use_container_width=True)
                
            mostrar_tabela_formatada(df.sort_values(['Grupo', 'Gap' if 'Gap' in df.columns else 'Grupo'], ascending=[True, False]), "Grupos.xlsx")

    with t4:
        df = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Conc_C_Sub" if dep_prefix=="VE" else f"PV_Tot_Conc_C_Sub"))
        if not df.empty:
            c1, c2, c3 = st.columns(3)
            with c1: r_sel = st.multiselect("Região", limpar(df['Região'].unique() if 'Região' in df.columns else []), key="reg_c")
            with c2: g_sel = st.multiselect("Grupo", limpar(df['Grupo'].unique() if 'Grupo' in df.columns else []), key="grp_c")
            with c3: l_sel = st.multiselect("Loja", limpar(df['Concessionária'].unique()), key="loja_c")
            if r_sel and 'Região' in df.columns: df = df[df['Região'].isin(r_sel)]
            if g_sel and 'Grupo' in df.columns: df = df[df['Grupo'].isin(g_sel)]
            if l_sel: df = df[df['Concessionária'].isin(l_sel)]
            
            exibir_tamanho_amostra(df)
            df_chart = df[~df['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)].copy()
            if not df_chart.empty and 'Gap' in df_chart.columns:
                st.subheader("Top 10 / Bottom 10 Impactos por Concessionária")
                df_chart['Gap'] = pd.to_numeric(df_chart['Gap'], errors='coerce').fillna(0.0)
                df_chart_agg = df_chart.groupby('Concessionária')['Gap'].sum().reset_index()
                df_chart_agg = df_chart_agg.rename(columns={'Gap': 'Impacto'})
                
                fig = gerar_grafico_impacto_corrigido(get_top_bottom_10(df_chart_agg, 'Impacto'), 'Concessionária')
                st.plotly_chart(fig, use_container_width=True)
                
            mostrar_tabela_formatada(df.sort_values(['Concessionária', 'Gap' if 'Gap' in df.columns else 'Concessionária'], ascending=[True, False]), "Concessionarias.xlsx")

    with t5:
        if dep_prefix == "VE":
            df = aplicar_filtros_globais(ler_dados_nps_oficial(file, "VE_Mod_C_Sub"))
            if not df.empty:
                sel = st.multiselect("Filtrar Modelo", limpar(df['Modelo'].unique()), key="mod_t5")
                if sel: df = df[df['Modelo'].isin(sel)]
                
                exibir_tamanho_amostra(df)
                df_chart = df[~df['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)].copy()
                if not df_chart.empty and 'Gap' in df_chart.columns:
                    st.subheader("Top 10 / Bottom 10 Impactos por Modelo")
                    df_chart['Gap'] = pd.to_numeric(df_chart['Gap'], errors='coerce').fillna(0.0)
                    df_chart_agg = df_chart.groupby('Modelo')['Gap'].sum().reset_index()
                    df_chart_agg = df_chart_agg.rename(columns={'Gap': 'Impacto'})
                    
                    fig = gerar_grafico_impacto_corrigido(get_top_bottom_10(df_chart_agg, 'Impacto'), 'Modelo')
                    st.plotly_chart(fig, use_container_width=True)
                    
                mostrar_tabela_formatada(df.sort_values(['Modelo', 'Gap' if 'Gap' in df.columns else 'Modelo'], ascending=[True, False]), "Modelos.xlsx")
        else:
            st.info("A visão por modelo no Pós-Vendas está consolidada na análise de 'Ciclo de Revisões'.")

# --- ANÁLISE DE NEUTROS / DETRATORES ---
elif tipo_analise in ["Análise de Neutros", "Análise de Detratores"]:
    foco = "Neutros" if "Neutros" in tipo_analise else "Detratores"
    file = f"Analise_{foco}_Yamaha.xlsx"
    st.title(f"⚖️ Potencial de Reversão: {foco} ({departamento})")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Regiões", "Grupos", "Concessionárias", "Causas Detalhadas", "Modelos"])
    
    with tab1:
        df_reg = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Potencial_Regiao"))
        if not df_reg.empty:
            filtro_local = st.multiselect("Filtrar Região", limpar(df_reg['Região'].unique()), key="sel_reg2")
            if filtro_local: df_reg = df_reg[df_reg['Região'].isin(filtro_local)]
            
            exibir_tamanho_amostra(df_reg)
            st.subheader("Ganho de NPS por Região")
            if 'Ganho Possível' in df_reg.columns:
                df_reg['Ganho Possível'] = pd.to_numeric(df_reg['Ganho Possível'], errors='coerce').fillna(0.0)
                fig = px.bar(df_reg, x='Região', y='Ganho Possível', color='Ganho Possível', text_auto='.1f')
                st.plotly_chart(fig, use_container_width=True)
            mostrar_tabela_formatada(df_reg, f"{foco}_Potencial_Regiao.xlsx")

    with tab2:
        df_grup = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Potencial_Grupo"))
        if not df_grup.empty:
            filtro_local = st.multiselect("Filtrar Grupo", limpar(df_grup['Grupo'].unique()), key="sel_grup2")
            if filtro_local: df_grup = df_grup[df_grup['Grupo'].isin(filtro_local)]
            
            exibir_tamanho_amostra(df_grup)
            st.subheader("Visão por Grupo")
            mostrar_tabela_formatada(df_grup, f"{foco}_Potencial_Grupo.xlsx")

    with tab3:
        df_conc = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Potencial_Conc"))
        if not df_conc.empty:
            filtro_local = st.multiselect("Filtrar Concessionária", limpar(df_conc['Concessionária'].unique()), key="sel_conc2")
            if filtro_local: df_conc = df_conc[df_conc['Concessionária'].isin(filtro_local)]
            
            exibir_tamanho_amostra(df_conc)
            st.subheader("Visão por Concessionária")
            mostrar_tabela_formatada(df_conc, f"{foco}_Potencial_Conc.xlsx")

    with tab4:
        df_causa = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Causas_{foco[:6]}"))
        if not df_causa.empty:
            segmentos = sorted(df_causa['Segmento_NPS'].dropna().unique())
            segmento = st.selectbox("Filtrar Tipo de Cliente", segmentos)
            if segmento: df_causa = df_causa[df_causa['Segmento_NPS'] == segmento]
            
            exibir_tamanho_amostra(df_causa)
            mostrar_tabela_formatada(df_causa, f"{foco}_Causas.xlsx")

    with tab5:
        if dep_prefix == "VE":
            df_mod = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"VE_Modelos_{foco[:6]}"))
            if not df_mod.empty:
                filtro_local = st.multiselect("Filtrar Modelo", limpar(df_mod['Modelo'].unique()), key="sel_mod2")
                if filtro_local: df_mod = df_mod[df_mod['Modelo'].isin(filtro_local)]
                
                exibir_tamanho_amostra(df_mod)
                st.subheader("Visão por Modelo")
                mostrar_tabela_formatada(df_mod, f"{foco}_Modelos.xlsx")
        else:
            st.info("A visão por modelo no Pós-Vendas está consolidada na análise de 'Ciclo de Revisões'.")

# --- CICLO DE REVISÕES ---
elif tipo_analise == "Ciclo de Revisões":
    if dep_prefix == "VE":
        st.title("🔧 Ciclo de Revisões")
        st.warning("Selecione 'Pós-Vendas' no menu lateral para visualizar esta análise.")
    else:
        df_cons = aplicar_filtros_globais(ler_dados_nps_oficial("Analise_Revisoes_Yamaha.xlsx", "Consolidado_PBI_Revisoes"))
        st.title("🔧 Ciclo de Revisões (Pós-Vendas)")
        
        if not df_cons.empty:
            if 'Modelo' in df_cons.columns:
                filtro_local = st.multiselect("Filtrar Modelo", limpar(df_cons['Modelo'].unique()), key="sel_mod_rev")
                if filtro_local: df_cons = df_cons[df_cons['Modelo'].isin(filtro_local)]

            if not df_cons.empty:
                exibir_tamanho_amostra(df_cons)
                
                df_cons['NPS'] = pd.to_numeric(df_cons['NPS'], errors='coerce').fillna(0.0)
                nps_evol = df_cons[~df_cons['Subcausa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')].groupby('Ciclo')['NPS'].mean().reset_index()
                
                ordem = ['Revisão 1', 'Revisão 2', 'Revisão 3', 'Revisão 4', 'Revisão 5 ou +']
                nps_evol['Ciclo'] = pd.Categorical(nps_evol['Ciclo'], categories=ordem, ordered=True)
                nps_evol = nps_evol.sort_values('Ciclo')
                
                if 'NPS' in nps_evol.columns:
                    fig = px.line(nps_evol, x='Ciclo', y='NPS', markers=True, text=[f"{x:.1f}" for x in nps_evol['NPS']], title="Evolução do NPS Médio por Ciclo")
                    fig.update_traces(textposition="top center")
                    st.plotly_chart(fig, use_container_width=True)
                
                rev_sel = st.selectbox("Detalhar Revisão", ordem)
                df_det = df_cons[df_cons['Ciclo'] == rev_sel]
                mostrar_tabela_formatada(df_det.drop(columns=['Ciclo', 'Aba_Origem'], errors='ignore'), "Detalhe_Revisoes.xlsx")
            else:
                st.info("Não há dados de Revisões para os filtros selecionados.")

# ==========================================
# 6. FOOTER / INFO DA MARCA
# ==========================================
st.sidebar.markdown("---")
st.sidebar.markdown(
    f'<div style="text-align: center; font-family: Montserrat;">'
    f'<span style="color:white; font-weight:700; font-size:18px;">INSIGHTS</span>'
    f'<span style="color:#FF6B6B; font-weight:900; font-size:20px;">&</span>'
    f'<span style="color:#9CA3AF; font-weight:400; font-size:16px;">Etc</span>'
    f'</div>', 
    unsafe_allow_html=True
)
st.sidebar.caption("<div style='text-align: center;'>Dashboard Seguro: Yamaha Motors</div>", unsafe_allow_html=True)