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

COR_AZUL = "#0A192F"
COR_TURQUESA = "#00D2D3"
COR_LARANJA = "#FF6B6B"

CORES_SEGMENTOS = {
    'Detratores': '#E63946',  
    'Neutros-': '#F4A261',    
    'Neutros+': '#E9C46A',    
    'Promotores': '#2A9D8F'   
}

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;900&display=swap');
    html, body, [class*="css"] {{ font-family: 'Montserrat', sans-serif; }}
    h1, h2, h3 {{ font-family: 'Montserrat', sans-serif !important; font-weight: 700 !important; color: {COR_AZUL}; }}
    section[data-testid="stSidebar"] {{ background-color: {COR_AZUL}; color: white; }}
    section[data-testid="stSidebar"] .stMarkdown p {{ color: white; }}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CONEXÃO COM BANCO DE DADOS (SQL SERVER)
# ==========================================
# Ativar quando a base estiver no Azure SQL (Configurar .streamlit/secrets.toml)
# try:
#     conn = st.connection('sqlserver', type='sql')
# except Exception as e:
#     st.warning("Conexão ao Azure SQL Server não estabelecida. Lendo em fallback mode (Arquivos Locais).")
#     conn = None

# ==========================================
# FUNÇÕES DE GRÁFICOS E UTILITÁRIOS
# ==========================================
def gerar_grafico_impacto_corrigido(df_plot, col_y, altura=350):
    if df_plot.empty or 'Impacto' not in df_plot.columns: return go.Figure().update_layout(title="Sem dados", plot_bgcolor='rgba(0,0,0,0)')
    df_plot = df_plot.copy()
    df_plot['Impacto'] = pd.to_numeric(df_plot['Impacto'], errors='coerce').fillna(0.0)
    df_plot = df_plot.sort_values('Impacto', ascending=True)
    eixo_y = df_plot[col_y].astype(str).tolist()
    valores_x = df_plot['Impacto'].tolist()
    cores = [COR_LARANJA if val < 0 else COR_TURQUESA for val in valores_x]
    textos = [f"{val:.3f}" if abs(val) > 0 and abs(val) < 0.1 else f"{val:.1f}" for val in valores_x]
    fig = go.Figure(go.Bar(x=valores_x, y=eixo_y, orientation='h', marker_color=cores, text=textos, textposition='outside', cliponaxis=False, textfont=dict(size=11, family='Montserrat')))
    max_abs = max([abs(v) for v in valores_x]) if valores_x else 1.0
    fig.update_layout(xaxis=dict(range=[-max_abs*1.3, max_abs*1.3], zeroline=True, zerolinewidth=2, zerolinecolor='rgba(0,0,0,0.3)', showgrid=False), yaxis=dict(tickfont=dict(size=11)), plot_bgcolor='rgba(0,0,0,0)', margin=dict(l=10, r=40, t=30, b=30), height=max(altura, len(valores_x) * 22))
    return fig

def gerar_matriz_dispersao(df, col_dimensao, altura=350):
    if df.empty or 'Gap' not in df.columns or 'NPS' not in df.columns: return go.Figure().update_layout(title="Sem dados para a Matriz", plot_bgcolor='rgba(0,0,0,0)')
    df_plot = df.copy()
    df_plot['Gap'] = pd.to_numeric(df_plot['Gap'], errors='coerce').fillna(0.0)
    df_plot['NPS'] = pd.to_numeric(df_plot['NPS'], errors='coerce').fillna(0.0)
    valores_x = df_plot['Gap'].tolist()
    valores_y = df_plot['NPS'].tolist()
    textos = df_plot[col_dimensao].astype(str).tolist()
    cores = [COR_LARANJA if val < 0 else COR_TURQUESA for val in valores_x]
    fig = go.Figure(go.Scatter(x=valores_x, y=valores_y, mode='markers+text', text=textos, textposition='top center', cliponaxis=False, marker=dict(size=12, color=cores, line=dict(width=1, color='DarkSlateGrey'))))
    fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(title='Impacto no NPS Nacional', zeroline=True, zerolinewidth=2, zerolinecolor='rgba(0,0,0,0.3)', showgrid=True, gridcolor='#E5E7EB'), yaxis=dict(title=f'NPS da {col_dimensao}', showgrid=True, gridcolor='#E5E7EB'), margin=dict(l=40, r=40, t=30, b=30), height=altura)
    return fig

def gerar_grafico_nps_barras(df, col_x, titulo="NPS"):
    if df.empty or 'NPS' not in df.columns: return go.Figure()
    df_plot = df.copy()
    df_plot['NPS'] = pd.to_numeric(df_plot['NPS'], errors='coerce').fillna(0.0)
    df_plot = df_plot.sort_values('NPS', ascending=False)
    x_vals = df_plot[col_x].astype(str).tolist()
    y_vals = df_plot['NPS'].tolist()
    cores = [COR_TURQUESA if v >= 0 else COR_LARANJA for v in y_vals]
    textos = [f"{v:.1f}" for v in y_vals]
    fig = go.Figure(go.Bar(x=x_vals, y=y_vals, marker_color=cores, text=textos, textposition='outside', cliponaxis=False, textfont=dict(size=11, family='Montserrat')))
    fig.update_layout(title=titulo, plot_bgcolor='rgba(0,0,0,0)', yaxis=dict(zeroline=True, zerolinewidth=2, zerolinecolor='gray', showgrid=True, gridcolor='#E5E7EB'), xaxis=dict(tickangle=-45), height=400, margin=dict(t=50, b=100))
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
    fig = go.Figure(go.Bar(x=eixo_x, y=valores_y, marker_color=cores, text=textos, textposition='outside', cliponaxis=False, textfont=dict(size=11, family='Montserrat')))
    max_abs = max([abs(v) for v in valores_y]) if valores_y else 1.0
    fig.update_layout(title=f"Impacto por {col_dimensao}: {causa}", plot_bgcolor='rgba(0,0,0,0)', yaxis=dict(range=[-max_abs*1.3, max_abs*1.3], zeroline=True, zerolinewidth=2, zerolinecolor='rgba(0,0,0,0.3)', showgrid=True, gridcolor='#E5E7EB'), xaxis=dict(tickangle=-45), height=400, margin=dict(t=50, b=100))
    return fig

def gerar_grafico_distribuicao_segmentos(df, col_dimensao):
    if df.empty: return go.Figure()
    df_plot = df.copy()
    vol_col = next((c for c in ['N_valido', 'Respondentes', 'Volume', 'Volume (N)'] if c in df_plot.columns), None)
    if vol_col: df_plot = df_plot.sort_values(vol_col, ascending=True)
    col_mapping = {}
    for col in df_plot.columns:
        c_norm = col.upper().replace(" ", "").replace('−', '-').replace('–', '-').replace('—', '-')
        if 'DETRATOR' in c_norm: col_mapping['Detratores'] = col
        elif 'NEUTRO-' in c_norm: col_mapping['Neutros-'] = col
        elif 'NEUTRO+' in c_norm: col_mapping['Neutros+'] = col
        elif 'PROMOTOR' in c_norm: col_mapping['Promotores'] = col
    fig = go.Figure()
    ordem_desejada = ['Detratores', 'Neutros-', 'Neutros+', 'Promotores']
    for label in ordem_desejada:
        if label in col_mapping:
            col_excel = col_mapping[label]
            x_vals = pd.to_numeric(df_plot[col_excel], errors='coerce').fillna(0).tolist()
            y_vals = df_plot[col_dimensao].astype(str).tolist()
            fig.add_trace(go.Bar(name=label, y=y_vals, x=x_vals, orientation='h', marker_color=CORES_SEGMENTOS[label], text=[f"{v:.0f}%" if v >= 5 else "" for v in x_vals], textposition='inside'))
    fig.update_layout(barmode='stack', plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(title="Distribuição %", range=[0, 100], showgrid=False, zeroline=False), yaxis=dict(tickfont=dict(size=11, family='Montserrat')), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), margin=dict(l=10, r=40, t=50, b=20), height=max(350, len(df_plot) * 35))
    return fig

# ==========================================
# 1. LOGIN E ACESSO
# ==========================================
USUARIOS = {
    "diretoria_yamaha": {"nome": "Diretoria Yamaha", "senha": "yamaha_nps_2026"},
    "Thyago": {"nome": "Thyago Angelo", "senha": "yamaha_nps_2026"},
    "admin": {"nome": "Fernando Deotti", "senha": "root_specialist"}
}
if "autenticado" not in st.session_state: st.session_state["autenticado"] = False

def verificar_login():
    u, s = st.session_state["campo_usuario"], st.session_state["campo_senha"]
    if u in USUARIOS and USUARIOS[u]["senha"] == s:
        st.session_state.update({"autenticado": True, "nome_usuario": USUARIOS[u]["nome"], "erro_login": False})
    else: st.session_state["erro_login"] = True

if not st.session_state["autenticado"]:
    c1, c2 = st.columns([1, 15])
    try: c1.image("image/LogoRoute.png", width=100)
    except: c1.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Yamaha_Motor_Logo.svg/1024px-Yamaha_Motor_Logo.svg.png", width=70)
    c2.title("Yamaha NPS Explorer")
    st.markdown("---")
    with st.columns([1, 2])[0]:
        st.subheader("Acesso Restrito")
        st.text_input("Usuário", key="campo_usuario")
        st.text_input("Senha", type="password", key="campo_senha")
        st.button("Entrar", on_click=verificar_login)
        if st.session_state.get("erro_login"): st.error("Credenciais inválidas.")
        st.markdown("<br><br>", unsafe_allow_html=True)
        # Identidade de Marca Insights&Etc Documentada
        st.markdown(f'<div style="text-align: left; font-family: Montserrat;"><span style="font-weight:700; font-size:16px;">INSIGHTS</span><span style="color:{COR_LARANJA}; font-weight:900; font-size:18px;">&</span><span style="font-weight:400; font-size:14px;">Etc</span></div>', unsafe_allow_html=True)
    st.stop()

# ==========================================
# 2. DADOS E BRITADEIRA
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
def ler_dados_nps_oficial(file_name, sheet_name):
    # [TODO_MIGRATION] -> Quando a Azure for liberada, comente a leitura do Excel e substitua por:
    # if conn is not None:
    #     try:
    #         query = f"SELECT * FROM [dbo].[{sheet_name}]"
    #         df = conn.query(query)
    #         return df
    #     except Exception as e:
    #         st.error(f"Erro ao ler banco SQL: {e}")
    #         return pd.DataFrame()
    
    path = os.path.join(OUTPUT_DIR, file_name)
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, engine='openpyxl')
            df.columns = df.columns.str.strip() 
            for col in df.columns:
                if col == 'Segmento_NPS': continue
                if any(term in col for term in ['NPS', 'Gap', 'Impacto', 'Contribuição', 'N_valido', 'Respondentes', 'Volume', '%', 'Ganho Possível']):
                    df[col] = df[col].apply(extrair_numero)
            return df
        except: return pd.DataFrame()
    return pd.DataFrame()

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

def exibir_tamanho_amostra(df):
    if df.empty: return
    n = next((df[c].sum() for c in ['N_valido', 'Respondentes', 'Volume (N)', 'Volume'] if c in df.columns), 0)
    st.markdown(f"<div style='text-align: right; font-size: 13px; color: #888; margin-bottom: 5px; font-family: Montserrat;'>Tamanho da Amostra (N): <b>{int(n):,}</b></div>".replace(',', '.'), unsafe_allow_html=True)

def mostrar_tabela_formatada(df, filename_download="dados.xlsx", hide_cols=None):
    if df.empty: return st.info("Sem dados.")
    df_disp = df.copy()
    if hide_cols: df_disp = df_disp.drop(columns=hide_cols, errors='ignore')
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

    fmt = {c: lambda x: f"{x:.1f}" if pd.notna(x) else "-" for c in df_disp.columns if any(t in c for t in ['NPS', 'Impacto', 'Contribuição', '% Respondentes', 'NPS Atual', 'NPS Potencial', 'Ganho Possível']) and c != 'Segmento_NPS'}
    if 'Respondentes' in df_disp.columns: fmt['Respondentes'] = lambda x: f"{int(x)}" if pd.notna(x) else "-"
    st.dataframe(df_disp.style.format(fmt), use_container_width=True, hide_index=True)
    st.download_button("📥 Baixar Excel", convert_df_to_excel(df_disp), filename_download, "application/vnd.ms-excel")

def aplicar_filtros_globais(df, f_causa, f_sub, f_nps, f_imp):
    if df.empty: return df
    if 'Região' in df.columns: df = df[~df['Região'].astype(str).str.upper().str.contains('PERFIL', na=False)]
    if f_causa and 'Causa da nota de recomendação' in df.columns: df = df[df['Causa da nota de recomendação'].isin(f_causa)]
    if f_sub and 'Subcausa da nota de recomendação' in df.columns: df = df[df['Subcausa da nota de recomendação'].isin(f_sub)]
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

# ==========================================
# 3. SIDEBAR E FILTROS GLOBAIS
# ==========================================
try:
    c_logo1, c_logo2, c_logo3 = st.sidebar.columns([1, 2, 1])
    with c_logo2: st.image("image/LogoRoute.png", use_container_width=True)
except: st.sidebar.title("ROUTE")

st.sidebar.markdown(f"<div style='text-align: center;'><b>Olá, {st.session_state['nome_usuario']}!</b></div>", unsafe_allow_html=True)
if st.sidebar.button("Sair (Logout)", use_container_width=True): st.session_state["autenticado"] = False; st.rerun()

st.sidebar.markdown("---")
departamento = st.sidebar.radio("Segmento", ["Vendas (VE)", "Pós-Vendas (PV)"])
dep_prefix = "VE" if departamento == "Vendas (VE)" else "PV"

# MENU PRINCIPAL
tipo_analise = st.sidebar.selectbox("Tipo de Análise", ["Resumo Executivo", "Contribuição Total", "Análise de Neutros", "Análise de Detratores", "Ciclo de Revisões"])

st.sidebar.markdown("---")
st.sidebar.subheader("Filtros Globais")
df_filtros = ler_dados_nps_oficial("Analise_NPS_Yamaha.xlsx", f"{dep_prefix}_Tot_C_Sub")

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

# --- RESUMO EXECUTIVO ---
if tipo_analise == "Resumo Executivo":
    st.title(f"📑 Resumo Executivo: {departamento}")
    st.markdown("Esta página consolida os principais insights, cruzando dados de contribuição, detratores, neutros e revisões.")
    
    df_causas_resumo = ler_dados_nps_oficial("Analise_NPS_Yamaha.xlsx", f"{dep_prefix}_Tot_Causa")
    df_subcausas_resumo = ler_dados_nps_oficial("Analise_NPS_Yamaha.xlsx", f"{dep_prefix}_Tot_C_Sub")
    df_reg_pot = ler_dados_nps_oficial("Analise_Neutros_Yamaha.xlsx", f"{dep_prefix}_Potencial_Regiao")
    df_reg_tot = ler_dados_nps_oficial("Analise_NPS_Yamaha.xlsx", f"{dep_prefix}_Reg_C_Sub" if dep_prefix=="VE" else "PV_Tot_Reg_C_Sub")
    df_rev = ler_dados_nps_oficial("Analise_Revisoes_Yamaha.xlsx", "Consolidado_PBI_Revisoes") if dep_prefix == "PV" else pd.DataFrame()

    if not df_causas_resumo.empty:
        try:
            # 1. Cálculo real do NPS e Componentes
            val_pro, val_neu, val_det = 0.0, 0.0, 0.0
            pct_neu_plus, pct_neu_minus, pct_det_plus, pct_det_minus = 0.0, 0.0, 0.0, 0.0
            
            if not df_reg_pot.empty:
                linha_tot = df_reg_pot[df_reg_pot['Região'].astype(str).str.upper().isin(['TOTAL NACIONAL', 'TOTAL'])]
                if not linha_tot.empty:
                    for c in linha_tot.columns:
                        c_upper = c.upper().replace(" ", "")
                        val = pd.to_numeric(linha_tot.iloc[0][c], errors='coerce')
                        if pd.isna(val): continue
                        if 'NEUTRO+' in c_upper: pct_neu_plus = val
                        elif 'NEUTRO-' in c_upper: pct_neu_minus = val
                        elif 'DETRATOR+' in c_upper: pct_det_plus = val
                        elif 'DETRATOR-' in c_upper: pct_det_minus = val
                        elif 'PROMOTOR' in c_upper: val_pro = val
                        elif 'NEUTRO' in c_upper and '+' not in c_upper and '-' not in c_upper: val_neu = val
                        elif 'DETRATOR' in c_upper and '+' not in c_upper and '-' not in c_upper: val_det = val

            if val_neu == 0.0 and (pct_neu_plus > 0 or pct_neu_minus > 0): val_neu = pct_neu_plus + pct_neu_minus
            if val_det == 0.0 and (pct_det_plus > 0 or pct_det_minus > 0): val_det = pct_det_plus + pct_det_minus
            
            df_causas_puras = df_causas_resumo[~df_causas_resumo['Causa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')]
            nps_atual = val_pro - val_det if val_pro > 0 else df_causas_puras['Contribuição'].sum()

            st.header("1. Panorama Geral do NPS")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("NPS Atual", f"{nps_atual:.1f}")
            c2.metric("Promotores", f"{val_pro:.1f}%")
            c3.metric("Neutros", f"{val_neu:.1f}%")
            c4.metric("Detratores", f"{val_det:.1f}%")

            st.header("2. Alavancas e Âncoras Principais")
            df_causas_obj = df_causas_resumo[~df_causas_resumo['Causa da nota de recomendação'].isin(termos_omitir_graficos + ['Outro(s) motivo(s)', 'Não indicou motivo específico'])]
            if not df_causas_obj.empty:
                df_causas_obj['Gap'] = pd.to_numeric(df_causas_obj['Gap'], errors='coerce').fillna(0.0)
                causa_top = df_causas_obj.loc[df_causas_obj['Gap'].idxmax()]
                causa_bot = df_causas_obj.loc[df_causas_obj['Gap'].idxmin()]
                st.markdown(f"**No nível de CAUSAS:**")
                st.markdown(f"* A principal causa objetiva que eleva o NPS é **{causa_top['Causa da nota de recomendação']}**, gerando um impacto positivo de **{causa_top['Gap']:.2f} pontos**.")
                st.markdown(f"* Por outro lado, a causa que mais prejudica a nota atual é **{causa_bot['Causa da nota de recomendação']}**, subtraindo **{causa_bot['Gap']:.2f} pontos**.")

            df_subcausas_obj = df_subcausas_resumo[~df_subcausas_resumo['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)]
            if not df_subcausas_obj.empty:
                df_subcausas_obj['Gap'] = pd.to_numeric(df_subcausas_obj['Gap'], errors='coerce').fillna(0.0)
                sub_top = df_subcausas_obj.loc[df_subcausas_obj['Gap'].idxmax()]
                sub_bot = df_subcausas_obj.loc[df_subcausas_obj['Gap'].idxmin()]
                st.markdown(f"**No nível de SUBCAUSAS:**")
                st.markdown(f"* A subcausa com melhor desempenho é **{sub_top['Subcausa da nota de recomendação']}** (Impacto de **{sub_top['Gap']:.2f} pontos**).")
                st.markdown(f"* O maior ofensor específico é **{sub_bot['Subcausa da nota de recomendação']}** (Impacto de **{sub_bot['Gap']:.2f} pontos**).")

            st.header("3. Oportunidades de Reversão (+ e -)")
            st.markdown(f"**Composição Qualificada:**")
            st.markdown(f"* Os Neutros dividem-se em **Neutros+ ({pct_neu_plus:.1f}%)** e Neutros- ({pct_neu_minus:.1f}%).")
            st.markdown(f"* Os Detratores dividem-se em **Detratores+ ({pct_det_plus:.1f}%)** e Detratores- ({pct_det_minus:.1f}%).")
            
            nps_pot_neu = nps_atual + pct_neu_plus
            nps_pot_det = nps_atual + (2 * pct_det_plus)
            nps_pot_total = nps_atual + pct_neu_plus + (2 * pct_det_plus)

            st.markdown("**Simulação de Impacto no NPS Geral:**")
            st.markdown(f"* Se convertermos todos os **Neutros+** em Promotores, o NPS subiria para **{nps_pot_neu:.1f}**.")
            st.markdown(f"* Se convertermos todos os **Detratores+** em Promotores, o NPS subiria para **{nps_pot_det:.1f}**.")
            st.markdown(f"* Se convertermos ambos (**Neutros+ e Detratores+**), o NPS dispararia para **{nps_pot_total:.1f}**.")

            st.header("4. Fortalezas e Fraquezas Geográficas")
            if not df_reg_tot.empty:
                df_reg_limpo = df_reg_tot[(df_reg_tot['Causa da nota de recomendação'].isin(['-', 'Total'])) & (df_reg_tot['N_valido'] >= 30)]
                if not df_reg_limpo.empty:
                    best_reg = df_reg_limpo.loc[df_reg_limpo['Gap'].idxmax()]
                    worst_reg = df_reg_limpo.loc[df_reg_limpo['Gap'].idxmin()]
                    st.markdown(f"**Fortaleza:** A **{best_reg['Região']}** é o maior motor do resultado, contribuindo com **{best_reg['Gap']:.2f} pontos** adicionais.")
                    st.markdown(f"**Fraqueza:** A **{worst_reg['Região']}** é a maior ofendora, derrubando o NPS em **{worst_reg['Gap']:.2f} pontos**.")
            
            if dep_prefix == "PV" and not df_rev.empty:
                st.header("5. Impacto do Ciclo de Revisões")
                df_rev_tot = df_rev[~df_rev['Subcausa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')].groupby('Ciclo').agg({'NPS':'mean', 'Gap':'sum', 'N_valido':'sum'}).reset_index()
                if not df_rev_tot.empty:
                    best_rev = df_rev_tot.loc[df_rev_tot['Gap'].idxmax()]
                    worst_rev = df_rev_tot.loc[df_rev_tot['Gap'].idxmin()]
                    nps_medio_rev = df_rev_tot['NPS'].mean()
                    st.markdown(f"A revisão que mais agrega valor é a **{best_rev['Ciclo']}** (NPS: {best_rev['NPS']:.1f}).")
                    st.markdown(f"A revisão com maior desgaste é a **{worst_rev['Ciclo']}** (NPS: {worst_rev['NPS']:.1f}).")
                    st.markdown(f"**Evolução Potencial:** Se a {worst_rev['Ciclo']} fosse nivelada à média do ciclo ({nps_medio_rev:.1f}), o NPS total de Pós-Vendas sofreria um ganho substancial de volume qualificado.")
        except Exception as e:
            st.warning(f"Erro ao compilar métricas do resumo executivo. Verifique a estrutura das planilhas.")
    else:
        st.info("Bases de dados não encontradas para gerar o Resumo Executivo.")

# --- CONTRIBUIÇÃO TOTAL ---
elif tipo_analise == "Contribuição Total":
    st.title(f"📊 Contribuição e Impactos: {departamento}")
    t1, t2, t3, t4, t5, t6 = st.tabs(["Geral", "Regiões", "Grupos", "Concessionárias", "Modelos", "Metodologia"])
    file = "Analise_NPS_Yamaha.xlsx"

    with t1:
        st.subheader("Filtros da Visão Geral")
        df_master = ler_dados_nps_oficial(file, f"{dep_prefix}_Master")
        df_c_default = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Tot_Causa"), f_causa, f_sub, f_nps, f_imp)
        df_s_default = aplicar_filtros_globais(ler_dados_nps_oficial(file, f"{dep_prefix}_Tot_C_Sub"), f_causa, f_sub, f_nps, f_imp)

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
                fig_wf = go.Figure(go.Bar(x=x_vals, y=y_vals, base=bases, marker_color=colors, text=[f"{v:.1f}" for v in y_vals], textposition='outside', textfont=dict(family='Montserrat')))
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

    def render_aba_comparativa(sheet_name, col_dimensao):
        df_raw = ler_dados_nps_oficial(file, sheet_name)
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

    # --- ABA METODOLOGIA RESTAURADA E COMPLETA ---
    with t6:
        st.header("Metodologia Insights&Etc")
        
        st.subheader("1. Como Calculamos o Impacto (Gap)")
        st.markdown("""
        O **Impacto (Gap)** revela se uma causa está puxando o NPS geral para cima (positivo) ou para baixo (negativo), ponderando a **representatividade** da causa no volume total.
        * **1. Contribuição Absoluta:** $Contribuição = \\left( \\frac{\\text{N da Causa}}{\\text{N Total}} \\right) \\times \\text{NPS da Causa}$
        * **2. Participação:** $Part = \\left( \\frac{\\text{Contribuição}}{\\text{NPS Total}} \\right) \\times 100$
        * **3. Peso:** $Peso = \\left( \\frac{\\text{N da Causa}}{\\text{N Total}} \\right) \\times 100$
        * **4. Impacto (Gap):** $Impacto = Part - Peso$
        
        ---
        
        ### Exemplo Prático
        Imagine que a Yamaha tem as seguintes métricas gerais num determinado período:
        * **NPS Total:** 50
        * **Respondentes Totais:** 1.000

        Analisando a causa **"Entrega Técnica"**, que possui:
        * **NPS da Entrega Técnica:** 80
        * **Respondentes da Entrega Técnica:** 200

        **Aplicando a fórmula:**
        1. **Contribuição:** $(200 / 1000) \\times 80 = 0.2 \\times 80 = \\mathbf{16}$ *(A Entrega contribui com 16 pontos absolutos para os 50 totais do NPS)*
        2. **Participação no NPS:** $(16 / 50) \\times 100 = \\mathbf{32\\%}$
        3. **Peso:** $(200 / 1000) \\times 100 = \\mathbf{20\\%}$
        4. **Impacto:** $32\\% - 20\\% = \\mathbf{+12}$

        **Interpretação:**
        A "Entrega Técnica" tem um **Impacto de +12**. Isso significa que, embora represente apenas 20% do volume de clientes (Peso), ela entrega 32% do resultado positivo do NPS. É uma **"alavanca"** que puxa a nota para cima.
        """)
        
        st.markdown("---")
        st.subheader("2. Quebras Estatísticas: O Poder do '+' e '-'")
        st.markdown("""
        Olhar apenas para o bloco inteiro esconde oportunidades de rápido retorno (Quick Wins). Por isso, dividimos as categorias:
        
        * **Promotores (Notas 9 e 10):** Clientes leais que recomendam a marca.
        * **Neutros+ (Nota 8):** Clientes quase promotores. Precisam de um leve encantamento para virar promotores.
        * **Neutros- (Nota 7):** Clientes em risco. Perto de caírem para a detração por frustrações silenciosas.
        * **Detratores+ (Notas 4 a 6):** Clientes insatisfeitos, mas recuperáveis com um plano de ação rápido.
        * **Detratores- (Notas 0 a 3):** Clientes extremamente frustrados. Custam muito mais esforço e recurso para reverter a percepção.
        
        **Estratégia:** Focar nas causas de impacto em **Neutros+** e **Detratores+** traz a maior e mais rápida alavancagem matemática para o NPS Global.
        """)

# --- ANÁLISE DE NEUTROS / DETRATORES ---
elif tipo_analise in ["Análise de Neutros", "Análise de Detratores"]:
    foco = "Neutros" if "Neutros" in tipo_analise else "Detratores"
    file = f"Analise_{foco}_Yamaha.xlsx"
    st.title(f"⚖️ Potencial de Reversão: {foco} ({departamento})")
    
    tabs = st.tabs(["Regiões", "Grupos", "Concessionárias", "Modelos", "Causas e Subcausas"])
    
    dims_configs = [
        ('Região', f"{dep_prefix}_Potencial_Regiao"),
        ('Grupo', f"{dep_prefix}_Potencial_Grupo"),
        ('Concessionária', f"{dep_prefix}_Potencial_Conc"),
        ('Modelo', f"VE_Modelos_{foco[:6]}" if dep_prefix == "VE" else "")
    ]
    
    for i, (dim, sheet) in enumerate(dims_configs):
        with tabs[i]: 
            if dim == 'Modelo' and dep_prefix == 'PV': 
                st.info("A visão por modelo no Pós-Vendas está consolidada na análise de 'Ciclo de Revisões'.")
                continue
            df_aba = aplicar_filtros_globais(ler_dados_nps_oficial(file, sheet), f_causa, f_sub, f_nps, f_imp)
            if not df_aba.empty:
                vol_col = next((c for c in ['N_valido', 'Respondentes', 'Volume', 'Volume (N)'] if c in df_aba.columns), None)
                if dim in ['Grupo', 'Concessionária', 'Modelo'] and vol_col:
                    if len(df_aba) > 10: df_aba = df_aba.sort_values(vol_col, ascending=False).head(10)
                exibir_tamanho_amostra(df_aba)
                st.subheader(f"Distribuição de Segmentos NPS por {dim}")
                st.plotly_chart(gerar_grafico_distribuicao_segmentos(df_aba, dim), use_container_width=True)
                st.subheader(f"Detalhamento Estatístico por {dim}")
                mostrar_tabela_formatada(df_aba, f"{foco}_{dim}.xlsx", hide_cols=['Ganho Possível', 'Potencial'])
            else: st.info(f"Nenhum dado encontrado para {foco} por {dim}.")

    with tabs[4]:
        st.subheader(f"Visão Geral de Impacto ({foco})")
        df_causas_nd = ler_dados_nps_oficial(file, f"{dep_prefix}_Causas_{foco[:6]}")
        if not df_causas_nd.empty:
            df_causas_nd = aplicar_filtros_globais(df_causas_nd, f_causa, f_sub, f_nps, f_imp)
            col1, col2, col3, col4 = st.columns(4)
            f_reg_nd, f_grp_nd, f_con_nd, f_mod_nd = [], [], [], []
            if 'Região' in df_causas_nd.columns:
                reg_vals = [r for r in df_causas_nd['Região'].dropna().unique() if 'PERFIL' not in str(r).upper()]
                f_reg_nd = col1.multiselect("Região", sorted(reg_vals), key=f"reg_nd_{foco}")
            if 'Grupo' in df_causas_nd.columns: f_grp_nd = col2.multiselect("Grupo", sorted(df_causas_nd['Grupo'].dropna().unique()), key=f"grp_nd_{foco}")
            if 'Concessionária' in df_causas_nd.columns: f_con_nd = col3.multiselect("Concessionária", sorted(df_causas_nd['Concessionária'].dropna().unique()), key=f"con_nd_{foco}")
            if 'Modelo' in df_causas_nd.columns: f_mod_nd = col4.multiselect("Modelo", sorted(df_causas_nd['Modelo'].dropna().unique()), key=f"mod_nd_{foco}")

            if f_reg_nd: df_causas_nd = df_causas_nd[df_causas_nd['Região'].isin(f_reg_nd)]
            if f_grp_nd: df_causas_nd = df_causas_nd[df_causas_nd['Grupo'].isin(f_grp_nd)]
            if f_con_nd: df_causas_nd = df_causas_nd[df_causas_nd['Concessionária'].isin(f_con_nd)]
            if f_mod_nd: df_causas_nd = df_causas_nd[df_causas_nd['Modelo'].isin(f_mod_nd)]
            
            if 'Segmento_NPS' in df_causas_nd.columns:
                segmentos = sorted(df_causas_nd['Segmento_NPS'].dropna().unique())
                segmento = st.selectbox("Filtrar Tipo de Cliente", ["Todos"] + list(segmentos), key=f"seg_nd_{foco}")
                if segmento != "Todos": df_causas_nd = df_causas_nd[df_causas_nd['Segmento_NPS'] == segmento]
            
            exibir_tamanho_amostra(df_causas_nd)
            st.markdown("---")
            if 'Gap' in df_causas_nd.columns and 'NPS' in df_causas_nd.columns:
                vol_col = next((c for c in ['N_valido', 'Respondentes', 'Volume', 'Volume (N)'] if c in df_causas_nd.columns), 'Volume')
                if vol_col not in df_causas_nd.columns: df_causas_nd[vol_col] = 1
                
                df_causas_nd['Gap'] = pd.to_numeric(df_causas_nd['Gap'], errors='coerce').fillna(0.0)
                df_causas_nd['NPS'] = pd.to_numeric(df_causas_nd['NPS'], errors='coerce').fillna(0.0)
                df_causas_nd[vol_col] = pd.to_numeric(df_causas_nd[vol_col], errors='coerce').fillna(0.0)
                
                df_c = df_causas_nd.groupby('Causa da nota de recomendação').agg({'Gap':'sum', 'NPS':'mean', vol_col:'sum'}).reset_index()
                df_s = df_causas_nd.groupby(['Causa da nota de recomendação', 'Subcausa da nota de recomendação']).agg({'Gap':'sum', 'NPS':'mean', vol_col:'sum'}).reset_index()
                
                # Trava N>=10 nas causas específicas
                df_c_plot = df_c[(~df_c['Causa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')) & (df_c[vol_col] >= 10)].copy()
                df_s_plot = df_s[(~df_s['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)) & (df_s[vol_col] >= 10)].copy()
                
                st.markdown("### Visão por Causa")
                col_a, col_b = st.columns(2)
                with col_a: st.plotly_chart(gerar_grafico_nps_barras(df_c_plot, 'Causa da nota de recomendação', 'NPS por Causa (N≥10)'), use_container_width=True)
                with col_b: st.plotly_chart(gerar_grafico_impacto_corrigido(df_c_plot.rename(columns={'Gap': 'Impacto'}), 'Causa da nota de recomendação', altura=400), use_container_width=True)
                
                st.markdown("### Visão por Subcausa")
                col_c, col_d = st.columns(2)
                with col_c: st.plotly_chart(gerar_grafico_nps_barras(df_s_plot, 'Subcausa da nota de recomendação', 'NPS por Subcausa (N≥10)'), use_container_width=True)
                with col_d: st.plotly_chart(gerar_grafico_impacto_corrigido(df_s_plot.rename(columns={'Gap': 'Impacto'}), 'Subcausa da nota de recomendação', altura=400), use_container_width=True)

            st.subheader("Tabelas de Detalhamento")
            mostrar_tabela_formatada(df_causas_nd, f"{foco}_Causas.xlsx", hide_cols=['Ganho Possível', 'Potencial'])
        else: st.info("Nenhum dado de causas detalhadas encontrado para este segmento.")

# --- CICLO DE REVISÕES COM SIMULAÇÃO ---
elif tipo_analise == "Ciclo de Revisões":
    if dep_prefix == "VE":
        st.title("🔧 Ciclo de Revisões")
        st.warning("Selecione 'Pós-Vendas' no menu lateral para visualizar esta análise.")
    else:
        file = "Analise_Revisoes_Yamaha.xlsx"
        df_cons = aplicar_filtros_globais(ler_dados_nps_oficial(file, "Consolidado_PBI_Revisoes"), f_causa, f_sub, f_nps, f_imp)
        st.title("🔧 Ciclo de Revisões (Pós-Vendas)")
        st.markdown("""
        Nesta análise, acompanhamos a jornada de manutenção do cliente. **É natural que o NPS apresente uma tendência de queda** à medida que as revisões avançam. Isso ocorre porque a motocicleta envelhece, os serviços tornam-se mais complexos e as peças fora da garantia encarecem o processo. O ponto de atenção é identificar **onde ocorre a queda mais brusca** para atuar de forma preventiva.
        """)
        
        if not df_cons.empty:
            # Remoção definitiva de 'Modelo' em PV para evitar distorções
            if 'Modelo' in df_cons.columns:
                df_cons = df_cons.drop(columns=['Modelo'])
            df_cons = df_cons.drop_duplicates()

            if not df_cons.empty:
                exibir_tamanho_amostra(df_cons)
                
                vol_col = next((c for c in ['N_valido', 'Respondentes', 'Volume (N)', 'Volume'] if c in df_cons.columns), 'Volume')
                df_cons['NPS'] = pd.to_numeric(df_cons['NPS'], errors='coerce').fillna(0.0)
                if vol_col not in df_cons.columns: df_cons[vol_col] = 1 
                else: df_cons[vol_col] = pd.to_numeric(df_cons[vol_col], errors='coerce').fillna(0.0)
                
                tab_evol, tab_conc = st.tabs(["Evolução e Simulação", "Desempenho por Concessionária"])
                
                with tab_evol:
                    nps_evol = df_cons[~df_cons['Subcausa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')].groupby('Ciclo').agg({'NPS':'mean', vol_col:'sum'}).reset_index()
                    ordem = ['Revisão 1', 'Revisão 2', 'Revisão 3', 'Revisão 4', 'Revisão 5 ou +']
                    nps_evol['Ciclo'] = pd.Categorical(nps_evol['Ciclo'], categories=ordem, ordered=True)
                    nps_evol = nps_evol.sort_values('Ciclo').reset_index(drop=True)
                    
                    if 'NPS' in nps_evol.columns and len(nps_evol) > 0:
                        nps_evol['Drop'] = nps_evol['NPS'].diff()
                        if len(nps_evol) > 1:
                            maior_queda_idx = nps_evol['Drop'].idxmin()
                            if pd.notna(maior_queda_idx):
                                st.info(f"💡 **Insight de Desgaste:** A maior queda na satisfação do cliente ocorre entre a **{nps_evol.loc[maior_queda_idx - 1, 'Ciclo']}** e a **{nps_evol.loc[maior_queda_idx, 'Ciclo']}**, com uma perda de **{abs(nps_evol.loc[maior_queda_idx, 'Drop']):.1f} pontos** de NPS.")

                        nps_medio_rev = (nps_evol['NPS'] * nps_evol[vol_col]).sum() / nps_evol[vol_col].sum() if nps_evol[vol_col].sum() > 0 else nps_evol['NPS'].mean()
                        
                        valores_x = nps_evol['Ciclo'].astype(str).tolist()
                        valores_y = [float(v) for v in nps_evol['NPS']]
                        textos = [f"{v:.1f}" for v in valores_y]
                        
                        fig = go.Figure(go.Scatter(
                            x=valores_x, y=valores_y, mode='lines+markers+text',
                            text=textos, textposition="top center",
                            line=dict(color=COR_TURQUESA, width=3),
                            marker=dict(size=10, color=COR_LARANJA),
                            textfont=dict(family='Montserrat')
                        ))
                        fig.add_hline(y=nps_medio_rev, line_dash="dot", line_color="gray", annotation_text=f"Média ({nps_medio_rev:.1f})")
                        fig.update_layout(title="Evolução do NPS Médio por Ciclo", plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor='#E5E7EB'), height=400, margin=dict(t=50, b=50))
                        st.plotly_chart(fig, use_container_width=True)
                        
                        st.markdown("---")
                        st.subheader("🎯 Simulação de Nivelamento (Potencial)")
                        st.markdown(f"Se as revisões com desempenho **abaixo da média ({nps_medio_rev:.1f})** fossem melhoradas para atingir exatamente a média do ciclo, qual seria o impacto global?")
                        
                        nps_evol['NPS_Simulado'] = nps_evol.apply(lambda row: nps_medio_rev if row['NPS'] < nps_medio_rev else row['NPS'], axis=1)
                        nps_rev_simulado = (nps_evol['NPS_Simulado'] * nps_evol[vol_col]).sum() / nps_evol[vol_col].sum() if nps_evol[vol_col].sum() > 0 else nps_evol['NPS_Simulado'].mean()
                        nps_rev_atual = (nps_evol['NPS'] * nps_evol[vol_col]).sum() / nps_evol[vol_col].sum() if nps_evol[vol_col].sum() > 0 else nps_evol['NPS'].mean()
                        
                        df_pv_tot = ler_dados_nps_oficial("Analise_NPS_Yamaha.xlsx", "PV_Tot_C_Sub")
                        nps_pv_atual, nps_pv_simulado = 0.0, 0.0
                        if not df_pv_tot.empty:
                            vol_pv_col = next((c for c in ['N_valido', 'Respondentes', 'Volume', 'Volume (N)'] if c in df_pv_tot.columns), None)
                            if vol_pv_col:
                                mask_tot = df_pv_tot['Causa da nota de recomendação'].astype(str).str.upper().isin(['-', 'TOTAL', 'TOTAL DA CAUSA'])
                                if mask_tot.any():
                                    n_pv_total = df_pv_tot[mask_tot][vol_pv_col].sum()
                                    nps_pv_atual = df_pv_tot[mask_tot]['NPS'].mean()
                                else:
                                    n_pv_total = df_pv_tot[~df_pv_tot['Causa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')][vol_pv_col].sum()
                                    if n_pv_total > 0: nps_pv_atual = (df_pv_tot[~df_pv_tot['Causa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')]['NPS'] * df_pv_tot[~df_pv_tot['Causa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')][vol_pv_col]).sum() / n_pv_total
                                
                                if n_pv_total > 0:
                                    gap_pontos_rev = ((nps_evol['NPS_Simulado'] - nps_evol['NPS']) * nps_evol[vol_col]).sum()
                                    nps_pv_simulado = nps_pv_atual + (gap_pontos_rev / n_pv_total)
                        
                        col_s1, col_s2 = st.columns(2)
                        col_s1.metric("NPS Total de Revisões", f"{nps_rev_simulado:.1f}", f"+{(nps_rev_simulado - nps_rev_atual):.1f} pontos vs. Atual ({nps_rev_atual:.1f})")
                        if nps_pv_atual != 0.0: col_s2.metric("NPS Total Pós-Vendas (PV)", f"{nps_pv_simulado:.1f}", f"+{(nps_pv_simulado - nps_pv_atual):.1f} pontos vs. Atual ({nps_pv_atual:.1f})")
                        else: col_s2.info("Base total de PV não encontrada para extrapolação.")
                    
                    st.markdown("---")
                    rev_sel = st.selectbox("Detalhar Revisão", ordem)
                    df_det = df_cons[df_cons['Ciclo'] == rev_sel]
                    df_det_view = df_det.drop(columns=['Ciclo', 'Aba_Origem'], errors='ignore').drop_duplicates()
                    mostrar_tabela_formatada(df_det_view, "Detalhe_Revisoes.xlsx")
                
                with tab_conc:
                    st.subheader("Ranking de Concessionárias nas Revisões")
                    if 'Concessionária' in df_cons.columns:
                        rev_sel_conc = st.selectbox("Selecione o Ciclo para o Ranking:", ['Todos'] + list(ordem), key='sel_rev_conc')
                        
                        if rev_sel_conc == 'Todos':
                            df_plot_conc = df_cons[~df_cons['Subcausa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')].groupby('Concessionária').agg({'NPS':'mean', vol_col:'sum'}).reset_index()
                        else:
                            df_plot_conc = df_cons[(df_cons['Ciclo'] == rev_sel_conc) & (~df_cons['Subcausa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL'))].groupby('Concessionária').agg({'NPS':'mean', vol_col:'sum'}).reset_index()
                            
                        # Trava de significância para ranking seguro
                        df_plot_conc = df_plot_conc[df_plot_conc[vol_col] >= 5]
                        
                        if not df_plot_conc.empty:
                            df_plot_conc = df_plot_conc.sort_values('NPS', ascending=False)
                            fig_conc = gerar_grafico_nps_barras(df_plot_conc, 'Concessionária', f"NPS por Concessionária ({rev_sel_conc}) - N≥5")
                            st.plotly_chart(fig_conc, use_container_width=True)
                            
                            st.dataframe(df_plot_conc.rename(columns={vol_col: 'Respondentes'}).style.format({'NPS': '{:.1f}'}), use_container_width=True, hide_index=True)
                        else:
                            st.info("Não há dados suficientes (N≥5) para gerar o ranking de concessionárias neste ciclo.")
                    else:
                        st.info("A dimensão 'Concessionária' não está presente nesta base de dados.")
            else: st.info("Não há dados de Revisões após aplicar os filtros.")
        else: st.info("Arquivo de Ciclo de Revisões não encontrado ou vazio.")

# ==========================================
# FOOTER / INFO DA MARCA
# ==========================================
st.sidebar.markdown("---")
# A assinatura Insights&Etc utilizando os guidelines da marca (Azul, Laranja e Turquesa)
st.sidebar.markdown(f'<div style="text-align: center; font-family: Montserrat;"><span style="color:white; font-weight:700; font-size:18px;">INSIGHTS</span><span style="color:{COR_LARANJA}; font-weight:900; font-size:20px;">&</span><span style="color:#9CA3AF; font-weight:400; font-size:16px;">Etc</span></div>', unsafe_allow_html=True)
st.sidebar.caption("<div style='text-align: center;'>Dashboard Seguro: Yamaha Motors</div>", unsafe_allow_html=True)