import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import re

# ==========================================
# IDENTIDADE VISUAL INSIGHTS&ETC
# ==========================================
st.set_page_config(page_title="Yamaha NPS Explorer", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Montserrat', sans-serif; }
    h1, h2, h3 { font-family: 'Montserrat', sans-serif !important; font-weight: 700 !important; }
    section[data-testid="stSidebar"] { background-color: #0A192F; color: white; }
    section[data-testid="stSidebar"] .stMarkdown p { color: white; }
    </style>
""", unsafe_allow_html=True)

COR_AZUL = "#0A192F"
COR_TURQUESA = "#00D2D3"
COR_LARANJA = "#FF6B6B"

# ==========================================
# FUNÇÕES DE GRÁFICOS (BLINDADOS GCP - LISTAS NATIVAS)
# ==========================================
def gerar_grafico_impacto_corrigido(df_plot, col_y, altura=350):
    if df_plot.empty or 'Impacto' not in df_plot.columns:
        return go.Figure().update_layout(title="Sem dados", plot_bgcolor='rgba(0,0,0,0)')

    df_plot = df_plot.copy()
    df_plot['Impacto'] = pd.to_numeric(df_plot['Impacto'], errors='coerce').fillna(0.0)
    df_plot = df_plot.sort_values('Impacto', ascending=True)

    eixo_y = df_plot[col_y].astype(str).tolist()
    valores_x = df_plot['Impacto'].tolist()

    cores = [COR_LARANJA if val < 0 else COR_TURQUESA for val in valores_x]
    textos = [f"{val:.3f}" if abs(val) > 0 and abs(val) < 0.1 else f"{val:.1f}" for val in valores_x]

    fig = go.Figure(go.Bar(
        x=valores_x, y=eixo_y, orientation='h', marker_color=cores,
        text=textos, textposition='outside', cliponaxis=False, textfont=dict(size=11)
    ))

    max_abs = max([abs(v) for v in valores_x]) if valores_x else 1.0
    max_range = (max_abs * 1.3) if max_abs > 0 else 1.0

    fig.update_layout(
        xaxis=dict(range=[-max_range, max_range], zeroline=True, zerolinewidth=2, zerolinecolor='rgba(0,0,0,0.3)', showgrid=False),
        yaxis=dict(tickfont=dict(size=11)),
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10, r=40, t=30, b=30),
        height=max(altura, len(valores_x) * 22)
    )
    return fig

def gerar_matriz_dispersao(df, col_dimensao, altura=350):
    if df.empty or 'Gap' not in df.columns or 'NPS' not in df.columns: 
        return go.Figure()
    
    df_plot = df.copy()
    df_plot['Gap'] = pd.to_numeric(df_plot['Gap'], errors='coerce').fillna(0.0)
    df_plot['NPS'] = pd.to_numeric(df_plot['NPS'], errors='coerce').fillna(0.0)

    # Conversão para listas nativas (Evita bug do Cloud Run)
    valores_x = df_plot['Gap'].tolist()
    valores_y = df_plot['NPS'].tolist()
    textos = df_plot[col_dimensao].astype(str).tolist()
    cores = [COR_LARANJA if val < 0 else COR_TURQUESA for val in valores_x]

    fig = go.Figure(go.Scatter(
        x=valores_x, y=valores_y, mode='markers+text',
        text=textos, textposition='top center',
        marker=dict(size=12, color=cores, line=dict(width=1, color='DarkSlateGrey'))
    ))

    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(title='Impacto no NPS Nacional', zeroline=True, zerolinewidth=2, zerolinecolor='rgba(0,0,0,0.3)', showgrid=True, gridcolor='#E5E7EB'),
        yaxis=dict(title=f'NPS da {col_dimensao}', showgrid=True, gridcolor='#E5E7EB'), 
        margin=dict(l=40, r=40, t=30, b=30),
        height=altura
    )
    return fig

def gerar_grafico_colunas_comparativo(df, col_dimensao, causa):
    if df.empty or 'Impacto' not in df.columns: return go.Figure()
    
    df_plot = df.copy()
    df_plot['Impacto'] = pd.to_numeric(df_plot['Impacto'], errors='coerce').fillna(0.0)
    df_plot = df_plot.sort_values('Impacto', ascending=False)
    
    eixo_x = df_plot[col_dimensao].astype(str).tolist()
    valores_y = df_plot['Impacto'].tolist()

    cores = [COR_LARANJA if val < 0 else COR_TURQUESA for val in valores_y]
    textos = [f"{val:.3f}" if abs(val) > 0 and abs(val) < 0.1 else f"{val:.1f}" for val in valores_y]

    fig = go.Figure(go.Bar(
        x=eixo_x, y=valores_y, marker_color=cores,
        text=textos, textposition='outside', cliponaxis=False, textfont=dict(size=11)
    ))

    max_abs = max([abs(v) for v in valores_y]) if valores_y else 1.0
    max_range = (max_abs * 1.3) if max_abs > 0 else 1.0

    fig.update_layout(
        title=f"Impacto por {col_dimensao}: {causa}",
        plot_bgcolor='rgba(0,0,0,0)',
        yaxis=dict(range=[-max_range, max_range], zeroline=True, zerolinewidth=2, zerolinecolor='rgba(0,0,0,0.3)', showgrid=True, gridcolor='#E5E7EB'),
        xaxis=dict(tickangle=-45), 
        height=400, margin=dict(t=50, b=100)
    )
    return fig

# ==========================================
# 1. LOGIN E CONFIGURAÇÕES BÁSICAS
# ==========================================
USUARIOS = {
    "diretoria_yamaha": {"nome": "Diretoria Yamaha", "senha": "yamaha_nps_2026"},
    "admin": {"nome": "Fernando Deotti", "senha": "root_specialist"}
}
if "autenticado" not in st.session_state: st.session_state["autenticado"] = False

def verificar_login():
    u, s = st.session_state["campo_usuario"], st.session_state["campo_senha"]
    if u in USUARIOS and USUARIOS[u]["senha"] == s:
        st.session_state.update({"autenticado": True, "nome_usuario": USUARIOS[u]["nome"], "erro_login": False})
    else:
        st.session_state["erro_login"] = True

def fazer_logout():
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    c1, c2 = st.columns([1, 15])
    try:
        c1.image("image/LogoRoute.png", width=100)
    except:
        c1.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Yamaha_Motor_Logo.svg/1024px-Yamaha_Motor_Logo.svg.png", width=70)
        
    c2.title("Yamaha NPS Explorer")
    st.markdown("---")
    with st.columns([1, 2])[0]:
        st.subheader("Acesso Restrito")
        st.text_input("Usuário", key="campo_usuario")
        st.text_input("Senha", type="password", key="campo_senha")
        st.button("Entrar", on_click=verificar_login)
        if st.session_state.get("erro_login"): st.error("Credenciais inválidas.")
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            f'<div style="text-align: left; font-family: Montserrat;">'
            f'<span style="font-weight:700; font-size:16px;">INSIGHTS</span>'
            f'<span style="color:#FF6B6B; font-weight:900; font-size:18px;">&</span>'
            f'<span style="font-weight:400; font-size:14px;">Etc</span>'
            f'</div>', unsafe_allow_html=True
        )
    st.stop()

# ==========================================
# 2. DADOS E BRITADEIRA (COM BUSCA INTELIGENTE DE ARQUIVOS)
# ==========================================
def extrair_numero(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    v = str(val).strip().replace('\xa0', '')
    if v.startswith('(') and v.endswith(')'): v = '-' + v[1:-1]
    v = re.sub(r'[^\d\.\-]', '', v.replace(',', '.').replace('−', '-').replace('–', '-').replace('—', '-'))
    try: return float(v) if v and v != '-' else 0.0
    except: return 0.0

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(os.path.dirname(BASE_DIR), 'output')

@st.cache_data
def ler_dados_nps_oficial(prefixo_arquivo, sheet_name):
    """ Busca o arquivo mais recente que comece com o prefixo informado (Resolve o problema do timestamp) """
    if not os.path.exists(OUTPUT_DIR): return pd.DataFrame()
    
    # Encontra todos os arquivos que começam com o prefixo
    arquivos_validos = [f for f in os.listdir(OUTPUT_DIR) if f.startswith(prefixo_arquivo) and f.endswith(('.xlsx', '.xls'))]
    if not arquivos_validos: return pd.DataFrame()
    
    # Pega o arquivo mais recente (útil se tiver vários Analise_Neutros_Data1, Data2, etc)
    arquivos_validos.sort(key=lambda x: os.path.getmtime(os.path.join(OUTPUT_DIR, x)), reverse=True)
    caminho_arquivo = os.path.join(OUTPUT_DIR, arquivos_validos[0])
    
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=sheet_name, engine='openpyxl')
        df.columns = df.columns.str.strip() 
        for col in ['NPS', 'Gap', 'Impacto', 'Contribuição', 'Peso', 'N_valido', 'Respondentes', 'Volume', 'Ganho Possível']:
            if col in df.columns: df[col] = df[col].apply(extrair_numero)
        return df
    except: return pd.DataFrame()

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

def exibir_tamanho_amostra(df):
    if df.empty: return
    n = 0
    for c in ['N_valido', 'Respondentes', 'Volume (N)', 'Volume']:
        if c in df.columns:
            n = df[c].sum()
            break
    st.markdown(f"<div style='text-align: right; font-size: 13px; color: #888; margin-bottom: 5px;'>Tamanho da Amostra (N): <b>{int(n):,}</b></div>".replace(',', '.'), unsafe_allow_html=True)

def mostrar_tabela_formatada(df, filename_download="dados.xlsx"):
    if df.empty: return st.info("Sem dados.")
    df_disp = df.copy()
    if 'Gap' in df_disp.columns and 'Impacto' in df_disp.columns: df_disp = df_disp.drop(columns=['Impacto'])
    df_disp = df_disp.rename(columns={'N_valido': 'Respondentes', 'Gap': 'Impacto'})
    
    lixo = ['-', '–', '—', 'Não especificada', 'Total', 'TOTAL DA CAUSA', 'nan', 'None', '']
    for col in ['Causa da nota de recomendação', 'Subcausa da nota de recomendação']:
        if col in df_disp.columns:
            df_disp[col] = df_disp[col].replace('Nenhuma das opções acima', 'Nenhuma subcausa específica')
            mask_valida = ~df_disp[col].astype(str).str.strip().isin(lixo)
            df_disp = df_disp[mask_valida]
    
    cols = list(df_disp.columns)
    if 'Respondentes' in cols and '% Respondentes' in cols:
        cols.remove('% Respondentes')
        cols.insert(cols.index('Respondentes') + 1, '% Respondentes')
        df_disp = df_disp[cols]

    fmt = {c: lambda x: f"{x:.1f}" if pd.notna(x) else "-" for c in ['NPS', 'Impacto', 'Contribuição', '% Respondentes', 'NPS Atual', 'NPS Potencial', 'Ganho Possível'] if c in df_disp.columns}
    if 'Respondentes' in df_disp.columns: fmt['Respondentes'] = lambda x: f"{int(x)}" if pd.notna(x) else "-"
    
    st.dataframe(df_disp.style.format(fmt), use_container_width=True, hide_index=True)
    st.download_button("📥 Baixar Excel", convert_df_to_excel(df_disp), filename_download, "application/vnd.ms-excel")

def aplicar_filtros_globais(df, f_causa, f_sub, f_nps, f_imp):
    if df.empty: return df
    if 'Região' in df.columns: df = df[~df['Região'].astype(str).str.upper().str.contains('PERFIL', na=False)]
    if f_causa and 'Causa da nota de recomendação' in df.columns: df = df[df['Causa da nota de recomendação'].isin(f_causa)]
    if f_sub and 'Subcausa da nota de recomendação' in df.columns: df = df[df['Subcausa da nota de recomendação'].isin(f_sub)]
    return df

# ==========================================
# 3. SIDEBAR E FILTROS GLOBAIS
# ==========================================
# LOGO REDUZIDO (Usa colunas para centralizar e diminuir o tamanho)
try:
    c_logo1, c_logo2, c_logo3 = st.sidebar.columns([1, 2, 1])
    with c_logo2:
        st.image("image/LogoRoute.png", use_container_width=True)
except:
    st.sidebar.title("ROUTE")

st.sidebar.markdown(f"<div style='text-align: center;'><b>Olá, {st.session_state['nome_usuario']}!</b></div>", unsafe_allow_html=True)
if st.sidebar.button("Sair (Logout)", use_container_width=True):
    st.session_state["autenticado"] = False
    st.rerun()

st.sidebar.markdown("---")
departamento = st.sidebar.radio("Segmento", ["Vendas (VE)", "Pós-Vendas (PV)"])
dep_prefix = "VE" if departamento == "Vendas (VE)" else "PV"

tipo_analise = st.sidebar.selectbox("Tipo de Análise", ["Contribuição Total", "Análise de Neutros", "Análise de Detratores", "Ciclo de Revisões"])

st.sidebar.markdown("---")
st.sidebar.subheader("Filtros Globais")
df_filtros = ler_dados_nps_oficial("Analise_NPS_Yamaha", f"{dep_prefix}_Tot_C_Sub")

def limpar(lista): return sorted([str(x) for x in lista if pd.notna(x) and str(x) not in ['-', 'Total', 'TOTAL DA CAUSA', 'Não especificada']])

if not df_filtros.empty:
    f_causa = st.sidebar.multiselect("Causa", limpar(df_filtros['Causa da nota de recomendação'].unique()))
    f_sub = st.sidebar.multiselect("Subcausa", limpar(df_filtros['Subcausa da nota de recomendação'].unique()))
else: f_causa, f_sub = [], []

f_nps = st.sidebar.multiselect("Faixa de NPS", ["Menor que 0", "0 a 49", "50 a 74", "75 a 100"])
f_imp = st.sidebar.multiselect("Impacto", ["Positivo", "Negativo"])

termos_omitir_graficos = ['-', 'Não especificada', 'Nenhuma das opções acima', 'Nenhuma subcausa específica', 'Total', 'TOTAL DA CAUSA']

# ==========================================
# 4. LÓGICA DE VISUALIZAÇÃO
# ==========================================
if tipo_analise == "Contribuição Total":
    st.title(f"📊 Contribuição e Impactos: {departamento}")
    t1, t2, t3, t4, t5 = st.tabs(["Geral", "Regiões", "Grupos", "Concessionárias", "Modelos"])
    file_prefix = "Analise_NPS_Yamaha"

    # --- ABA GERAL ---
    with t1:
        st.subheader("Filtros da Visão Geral")
        df_master = ler_dados_nps_oficial(file_prefix, f"{dep_prefix}_Master")
        df_c_default = aplicar_filtros_globais(ler_dados_nps_oficial(file_prefix, f"{dep_prefix}_Tot_Causa"), f_causa, f_sub, f_nps, f_imp)
        df_s_default = aplicar_filtros_globais(ler_dados_nps_oficial(file_prefix, f"{dep_prefix}_Tot_C_Sub"), f_causa, f_sub, f_nps, f_imp)

        col1, col2, col3, col4 = st.columns(4)
        f_reg, f_grp, f_con, f_mod = [], [], [], []
        
        if not df_master.empty:
            if 'Região' in df_master.columns: 
                reg_vals = [r for r in df_master['Região'].dropna().unique() if 'PERFIL' not in str(r).upper()]
                f_reg = col1.multiselect("Região", sorted(reg_vals))
            if 'Grupo' in df_master.columns: f_grp = col2.multiselect("Grupo", sorted(df_master['Grupo'].dropna().unique()))
            if 'Concessionária' in df_master.columns: f_con = col3.multiselect("Concessionária", sorted(df_master['Concessionária'].dropna().unique()))
            if 'Modelo' in df_master.columns: f_mod = col4.multiselect("Modelo", sorted(df_master['Modelo'].dropna().unique()))

        if (f_reg or f_grp or f_con or f_mod) and not df_master.empty:
            df_g = df_master.copy()
            if f_reg: df_g = df_g[df_g['Região'].isin(f_reg)]
            if f_grp: df_g = df_g[df_g['Grupo'].isin(f_grp)]
            if f_con: df_g = df_g[df_g['Concessionária'].isin(f_con)]
            if f_mod: df_g = df_g[df_g['Modelo'].isin(f_mod)]
            df_g = aplicar_filtros_globais(df_g, f_causa, f_sub, f_nps, f_imp)
            df_c = df_g.groupby('Causa da nota de recomendação').agg({'N_valido':'sum', 'NPS':'mean', 'Contribuição':'sum', 'Gap':'sum'}).reset_index()
            df_s = df_g.groupby(['Causa da nota de recomendação', 'Subcausa da nota de recomendação']).agg({'N_valido':'sum', 'NPS':'mean', 'Contribuição':'sum', 'Gap':'sum'}).reset_index()
        else:
            df_c, df_s = df_c_default, df_s_default

        exibir_tamanho_amostra(df_s)
        st.markdown("---")
        if not df_c.empty:
            st.subheader("Construção do NPS")
            df_wf = df_c[~df_c['Causa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')].copy()
            ordem_wf = ["Não indicou motivo específico", "Instalações físicas", "Equipe de consultores" if dep_prefix == "PV" else "Equipe de vendedores", "Momento da negociação", "Entrega técnica", "Outro(s) motivo(s)"]
            df_wf['Causa da nota de recomendação'] = pd.Categorical(df_wf['Causa da nota de recomendação'].replace(['-', 'Não especificada'], 'Não indicou motivo específico'), categories=ordem_wf, ordered=True)
            df_wf = df_wf.sort_values('Causa da nota de recomendação').dropna(subset=['Causa da nota de recomendação'])
            if not df_wf.empty:
                x_vals = list(df_wf['Causa da nota de recomendação']) + ["NPS Total"]
                deltas = list(df_wf['Contribuição'])
                y_vals = deltas + [sum(deltas)]
                bases, atual = [], 0
                for d in deltas: bases.append(atual); atual += d
                bases.append(0)
                colors = ["#1f77b4" if c=="NPS Total" else ("#808080" if "motivo" in c else (COR_TURQUESA if v>=0 else COR_LARANJA)) for c, v in zip(x_vals, y_vals)]
                fig_wf = go.Figure(go.Bar(x=x_vals, y=y_vals, base=bases, marker_color=colors, text=[f"{v:.1f}" for v in y_vals], textposition='outside'))
                fig_wf.update_layout(showlegend=False, plot_bgcolor='rgba(0,0,0,0)', height=400)
                st.plotly_chart(fig_wf, use_container_width=True)

        if not df_s.empty:
            st.subheader("Matriz de Impacto por Subcausa")
            df_gap = df_s[~df_s['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)].copy()
            if 'Gap' in df_gap.columns:
                df_gap = df_gap.rename(columns={'Gap': 'Impacto'})
                st.plotly_chart(gerar_grafico_impacto_corrigido(df_gap, 'Subcausa da nota de recomendação'), use_container_width=True)

        st.subheader("Impacto Consolidado por Causa")
        mostrar_tabela_formatada(df_c, "Geral_Causa.xlsx")
        st.subheader("Detalhe por Subcausa")
        mostrar_tabela_formatada(df_s, "Geral_Subcausa.xlsx")

    # --- ABAS COMPARATIVAS ---
    def render_aba_comparativa(sheet_name, col_dimensao):
        df_raw = ler_dados_nps_oficial(file_prefix, sheet_name)
        if df_raw.empty: return st.info(f"Nenhum dado encontrado para a visão de {col_dimensao}.")
        
        mask_totais_base = (df_raw['Causa da nota de recomendação'].isin(['-', 'Total'])) & (df_raw['Subcausa da nota de recomendação'].isin(['-', 'Total']))
        entidades_validas = df_raw[mask_totais_base & (df_raw['N_valido'] >= 30)][col_dimensao].unique()
        df_comp = df_raw[df_raw[col_dimensao].isin(entidades_validas)]
        
        if df_comp.empty: return st.info(f"Nenhum(a) {col_dimensao} possui amostra mínima de 30 entrevistas.")

        df_comp = aplicar_filtros_globais(df_comp, f_causa, f_sub, f_nps, f_imp)
        mask_totais = (df_comp['Causa da nota de recomendação'].isin(['-', 'Total'])) & (df_comp['Subcausa da nota de recomendação'].isin(['-', 'Total']))
        df_totais = df_comp[mask_totais].copy()

        if col_dimensao in ['Grupo', 'Concessionária'] and len(df_totais) > 20:
            df_sorted = df_totais.sort_values('Gap')
            foco = pd.concat([df_sorted.head(10), df_sorted.tail(10)])[col_dimensao].unique()
            df_totais = df_totais[df_totais[col_dimensao].isin(foco)]
            mask_causas = (~df_comp['Causa da nota de recomendação'].isin(['-', 'Total'])) & (df_comp['Subcausa da nota de recomendação'].isin(['-', 'Total'])) & (df_comp[col_dimensao].isin(foco))
        else:
            mask_causas = (~df_comp['Causa da nota de recomendação'].isin(['-', 'Total'])) & (df_comp['Subcausa da nota de recomendação'].isin(['-', 'Total']))
            
        df_causas = df_comp[mask_causas].copy()
        if 'Gap' in df_causas.columns: df_causas = df_causas.rename(columns={'Gap': 'Impacto'}) 

        if not df_totais.empty:
            c1, c2 = st.columns(2)
            alt = max(350, len(df_totais) * 22)
            with c1:
                st.subheader(f"Impacto Geral por {col_dimensao}")
                st.plotly_chart(gerar_grafico_impacto_corrigido(df_totais.rename(columns={'Gap': 'Impacto'}), col_dimensao, altura=alt), use_container_width=True)
            with c2:
                st.subheader("Matriz de Desempenho")
                st.plotly_chart(gerar_matriz_dispersao(df_totais, col_dimensao, altura=alt), use_container_width=True)

        if not df_causas.empty:
            st.markdown("---")
            st.subheader(f"Impacto Comparativo por Causa ({col_dimensao})")
            lista = sorted(df_causas['Causa da nota de recomendação'].unique())
            cols = st.columns(2)
            for i, causa in enumerate(lista):
                with cols[i % 2]:
                    st.plotly_chart(gerar_grafico_colunas_comparativo(df_causas[df_causas['Causa da nota de recomendação'] == causa], col_dimensao, causa), use_container_width=True)

    with t2: render_aba_comparativa(f"{dep_prefix}_Reg_C_Sub" if dep_prefix=="VE" else "PV_Tot_Reg_C_Sub", 'Região')
    with t3: render_aba_comparativa(f"{dep_prefix}_Grup_C_Sub" if dep_prefix=="VE" else "PV_Tot_Grup_C_Sub", 'Grupo')
    with t4: render_aba_comparativa(f"{dep_prefix}_Conc_C_Sub" if dep_prefix=="VE" else "PV_Tot_Conc_C_Sub", 'Concessionária')
    with t5: 
        if dep_prefix == "VE": render_aba_comparativa("VE_Mod_C_Sub", 'Modelo')
        else: st.info("Análise de modelo disponível em 'Ciclo de Revisões'.")

# --- OUTRAS ANÁLISES ---
elif tipo_analise in ["Análise de Neutros", "Análise de Detratores"]:
    foco = "Neutros" if "Neutros" in tipo_analise else "Detratores"
    file_prefix = f"Analise_{foco}_Yamaha"
    st.title(f"⚖️ Potencial de Reversão: {foco} ({departamento})")
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Regiões", "Grupos", "Concessionárias", "Causas Detalhadas", "Modelos"])
    
    with tab1:
        df_reg = aplicar_filtros_globais(ler_dados_nps_oficial(file_prefix, f"{dep_prefix}_Potencial_Regiao"), f_causa, f_sub, f_nps, f_imp)
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
        else: st.info(f"Nenhum dado encontrado para {foco} por Região.")

    with tab2:
        df_grup = aplicar_filtros_globais(ler_dados_nps_oficial(file_prefix, f"{dep_prefix}_Potencial_Grupo"), f_causa, f_sub, f_nps, f_imp)
        if not df_grup.empty:
            filtro_local = st.multiselect("Filtrar Grupo", limpar(df_grup['Grupo'].unique()), key="sel_grup2")
            if filtro_local: df_grup = df_grup[df_grup['Grupo'].isin(filtro_local)]
            exibir_tamanho_amostra(df_grup)
            st.subheader("Visão por Grupo")
            mostrar_tabela_formatada(df_grup, f"{foco}_Potencial_Grupo.xlsx")
        else: st.info(f"Nenhum dado encontrado para {foco} por Grupo.")

    with tab3:
        df_conc = aplicar_filtros_globais(ler_dados_nps_oficial(file_prefix, f"{dep_prefix}_Potencial_Conc"), f_causa, f_sub, f_nps, f_imp)
        if not df_conc.empty:
            filtro_local = st.multiselect("Filtrar Concessionária", limpar(df_conc['Concessionária'].unique()), key="sel_conc2")
            if filtro_local: df_conc = df_conc[df_conc['Concessionária'].isin(filtro_local)]
            exibir_tamanho_amostra(df_conc)
            st.subheader("Visão por Concessionária")
            mostrar_tabela_formatada(df_conc, f"{foco}_Potencial_Conc.xlsx")
        else: st.info(f"Nenhum dado encontrado para {foco} por Concessionária.")

    with tab4:
        df_causa_aba = aplicar_filtros_globais(ler_dados_nps_oficial(file_prefix, f"{dep_prefix}_Causas_{foco[:6]}"), f_causa, f_sub, f_nps, f_imp)
        if not df_causa_aba.empty:
            segmentos = sorted(df_causa_aba['Segmento_NPS'].dropna().unique())
            segmento = st.selectbox("Filtrar Tipo de Cliente", segmentos)
            if segmento: df_causa_aba = df_causa_aba[df_causa_aba['Segmento_NPS'] == segmento]
            exibir_tamanho_amostra(df_causa_aba)
            mostrar_tabela_formatada(df_causa_aba, f"{foco}_Causas.xlsx")
        else: st.info(f"Nenhum dado de causas detalhadas encontrado para {foco}.")

    with tab5:
        if dep_prefix == "VE":
            df_mod = aplicar_filtros_globais(ler_dados_nps_oficial(file_prefix, f"VE_Modelos_{foco[:6]}"), f_causa, f_sub, f_nps, f_imp)
            if not df_mod.empty:
                filtro_local = st.multiselect("Filtrar Modelo", limpar(df_mod['Modelo'].unique()), key="sel_mod2")
                if filtro_local: df_mod = df_mod[df_mod['Modelo'].isin(filtro_local)]
                exibir_tamanho_amostra(df_mod)
                st.subheader("Visão por Modelo")
                mostrar_tabela_formatada(df_mod, f"{foco}_Modelos.xlsx")
            else: st.info(f"Nenhum dado encontrado para {foco} por Modelo.")
        else: st.info("A visão por modelo no Pós-Vendas está consolidada na análise de 'Ciclo de Revisões'.")

elif tipo_analise == "Ciclo de Revisões":
    if dep_prefix == "VE":
        st.title("🔧 Ciclo de Revisões")
        st.warning("Selecione 'Pós-Vendas' no menu lateral para visualizar esta análise.")
    else:
        file_prefix = "Analise_Revisoes_Yamaha"
        df_cons = aplicar_filtros_globais(ler_dados_nps_oficial(file_prefix, "Consolidado_PBI_Revisoes"), f_causa, f_sub, f_nps, f_imp)
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
            else: st.info("Não há dados de Revisões após aplicar os filtros.")
        else:
            st.info("Arquivo de Ciclo de Revisões não encontrado ou vazio.")

# ==========================================
# FOOTER
# ==========================================
st.sidebar.markdown("---")
st.sidebar.markdown(
    f'<div style="text-align: center; font-family: Montserrat;">'
    f'<span style="color:white; font-weight:700; font-size:18px;">INSIGHTS</span>'
    f'<span style="color:{COR_LARANJA}; font-weight:900; font-size:20px;">&</span>'
    f'<span style="color:#9CA3AF; font-weight:400; font-size:16px;">Etc</span>'
    f'</div>', unsafe_allow_html=True
)
st.sidebar.caption("<div style='text-align: center;'>Dashboard Seguro: Yamaha Motors</div>", unsafe_allow_html=True)