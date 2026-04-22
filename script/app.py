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

CORES_SEGMENTOS = {
    'Detratores': '#E63946',  
    'Neutros-': '#F4A261',    
    'Neutros+': '#E9C46A',    
    'Promotores': '#2A9D8F'   
}

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
        return go.Figure().update_layout(title="Sem dados para a Matriz", plot_bgcolor='rgba(0,0,0,0)')
    
    df_plot = df.copy()
    df_plot['Gap'] = pd.to_numeric(df_plot['Gap'], errors='coerce').fillna(0.0)
    df_plot['NPS'] = pd.to_numeric(df_plot['NPS'], errors='coerce').fillna(0.0)

    valores_x = df_plot['Gap'].tolist()
    valores_y = df_plot['NPS'].tolist()
    textos = df_plot[col_dimensao].astype(str).tolist()
    cores = [COR_LARANJA if val < 0 else COR_TURQUESA for val in valores_x]

    fig = go.Figure(go.Scatter(
        x=valores_x, y=valores_y, mode='markers+text',
        text=textos, textposition='top center', cliponaxis=False,
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

def gerar_grafico_distribuicao_segmentos(df, col_dimensao):
    if df.empty: return go.Figure()
    df_plot = df.copy()
    
    vol_col = next((c for c in ['N_valido', 'Respondentes', 'Volume', 'Volume (N)'] if c in df_plot.columns), None)
    if vol_col:
        df_plot = df_plot.sort_values(vol_col, ascending=True)
    
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
            
            fig.add_trace(go.Bar(
                name=label,
                y=y_vals,
                x=x_vals,
                orientation='h',
                marker_color=CORES_SEGMENTOS[label],
                text=[f"{v:.0f}%" if v >= 5 else "" for v in x_vals], 
                textposition='inside'
            ))

    fig.update_layout(
        barmode='stack',
        plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(title="Distribuição %", range=[0, 100], showgrid=False, zeroline=False),
        yaxis=dict(tickfont=dict(size=11)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=10, r=40, t=50, b=20),
        height=max(350, len(df_plot) * 35)
    )
    return fig

# ==========================================
# 1. LOGIN E CONFIGURAÇÕES BÁSICAS
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

def mostrar_tabela_formatada(df, filename_download="dados.xlsx", hide_cols=None):
    if df.empty: return st.info("Sem dados.")
    df_disp = df.copy()
    
    if hide_cols: 
        df_disp = df_disp.drop(columns=hide_cols, errors='ignore')
        
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
if tipo_analise == "Contribuição Total":
    st.title(f"📊 Contribuição e Impactos: {departamento}")
    t1, t2, t3, t4, t5 = st.tabs(["Geral", "Regiões", "Grupos", "Concessionárias", "Modelos"])
    file = "Analise_NPS_Yamaha.xlsx"

    # --- ABA GERAL ---
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

# --- ANÁLISE DE NEUTROS / DETRATORES ---
elif tipo_analise in ["Análise de Neutros", "Análise de Detratores"]:
    foco = "Neutros" if "Neutros" in tipo_analise else "Detratores"
    file = f"Analise_{foco}_Yamaha.xlsx"
    st.title(f"⚖️ Potencial de Reversão: {foco} ({departamento})")
    
    # Aba "Causas e Subcausas" agora é a ÚLTIMA da lista
    tabs = st.tabs(["Regiões", "Grupos", "Concessionárias", "Modelos", "Causas e Subcausas"])
    
    # === 1. ABAS COMPARATIVAS (Região, Grupo, Conc, Modelo) ===
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
                    if len(df_aba) > 10:
                        df_aba = df_aba.sort_values(vol_col, ascending=False).head(10)
                
                exibir_tamanho_amostra(df_aba)
                
                st.subheader(f"Distribuição de Segmentos NPS por {dim}")
                st.plotly_chart(gerar_grafico_distribuicao_segmentos(df_aba, dim), use_container_width=True)
                
                st.subheader(f"Detalhamento Estatístico por {dim}")
                mostrar_tabela_formatada(df_aba, f"{foco}_{dim}.xlsx", hide_cols=['Ganho Possível', 'Potencial'])
            else:
                st.info(f"Nenhum dado encontrado para {foco} por {dim}.")

    # === 2. ABA CAUSAS E SUBCAUSAS (Agora no index 4) ===
    with tabs[4]:
        st.subheader(f"Visão Geral de Impacto ({foco})")
        df_causas_nd = ler_dados_nps_oficial(file, f"{dep_prefix}_Causas_{foco[:6]}")
        
        if not df_causas_nd.empty:
            df_causas_nd = aplicar_filtros_globais(df_causas_nd, f_causa, f_sub, f_nps, f_imp)
            
            # FILTROS EXCLUSIVOS DA ABA CAUSAS
            col1, col2, col3, col4 = st.columns(4)
            f_reg_nd, f_grp_nd, f_con_nd, f_mod_nd = [], [], [], []
            
            if 'Região' in df_causas_nd.columns:
                reg_vals = [r for r in df_causas_nd['Região'].dropna().unique() if 'PERFIL' not in str(r).upper()]
                f_reg_nd = col1.multiselect("Região", sorted(reg_vals), key=f"reg_nd_{foco}")
            if 'Grupo' in df_causas_nd.columns: 
                f_grp_nd = col2.multiselect("Grupo", sorted(df_causas_nd['Grupo'].dropna().unique()), key=f"grp_nd_{foco}")
            if 'Concessionária' in df_causas_nd.columns: 
                f_con_nd = col3.multiselect("Concessionária", sorted(df_causas_nd['Concessionária'].dropna().unique()), key=f"con_nd_{foco}")
            if 'Modelo' in df_causas_nd.columns: 
                f_mod_nd = col4.multiselect("Modelo", sorted(df_causas_nd['Modelo'].dropna().unique()), key=f"mod_nd_{foco}")

            # Aplicação dos filtros locais de dimensão
            if f_reg_nd: df_causas_nd = df_causas_nd[df_causas_nd['Região'].isin(f_reg_nd)]
            if f_grp_nd: df_causas_nd = df_causas_nd[df_causas_nd['Grupo'].isin(f_grp_nd)]
            if f_con_nd: df_causas_nd = df_causas_nd[df_causas_nd['Concessionária'].isin(f_con_nd)]
            if f_mod_nd: df_causas_nd = df_causas_nd[df_causas_nd['Modelo'].isin(f_mod_nd)]
            
            # Filtro de Segmento (se aplicável)
            if 'Segmento_NPS' in df_causas_nd.columns:
                segmentos = sorted(df_causas_nd['Segmento_NPS'].dropna().unique())
                segmento = st.selectbox("Filtrar Tipo de Cliente", ["Todos"] + list(segmentos), key=f"seg_nd_{foco}")
                if segmento != "Todos":
                    df_causas_nd = df_causas_nd[df_causas_nd['Segmento_NPS'] == segmento]
            
            exibir_tamanho_amostra(df_causas_nd)
            st.markdown("---")
            
            # Agregação dinâmica após os filtros
            if 'Gap' in df_causas_nd.columns:
                df_causas_nd['Gap'] = pd.to_numeric(df_causas_nd['Gap'], errors='coerce').fillna(0.0)
                
                df_c = df_causas_nd.groupby('Causa da nota de recomendação')[['Gap']].sum().reset_index()
                df_s = df_causas_nd.groupby(['Causa da nota de recomendação', 'Subcausa da nota de recomendação'])[['Gap']].sum().reset_index()
                
                col_a, col_b = st.columns(2)
                with col_a:
                    st.subheader("Impacto por Causa")
                    df_c_plot = df_c[~df_c['Causa da nota de recomendação'].astype(str).str.upper().str.startswith('TOTAL')].copy()
                    df_c_plot = df_c_plot.rename(columns={'Gap': 'Impacto'})
                    st.plotly_chart(gerar_grafico_impacto_corrigido(df_c_plot, 'Causa da nota de recomendação', altura=400), use_container_width=True)
                with col_b:
                    st.subheader("Impacto por Subcausa")
                    df_s_plot = df_s[~df_s['Subcausa da nota de recomendação'].isin(termos_omitir_graficos)].copy()
                    df_s_plot = df_s_plot.rename(columns={'Gap': 'Impacto'})
                    st.plotly_chart(gerar_grafico_impacto_corrigido(df_s_plot, 'Subcausa da nota de recomendação', altura=400), use_container_width=True)
            
            st.subheader("Tabelas de Detalhamento")
            mostrar_tabela_formatada(df_causas_nd, f"{foco}_Causas.xlsx", hide_cols=['Ganho Possível', 'Potencial'])
        else:
            st.info("Nenhum dado de causas detalhadas encontrado para este segmento.")

# --- CICLO DE REVISÕES ---
elif tipo_analise == "Ciclo de Revisões":
    if dep_prefix == "VE":
        st.title("🔧 Ciclo de Revisões")
        st.warning("Selecione 'Pós-Vendas' no menu lateral para visualizar esta análise.")
    else:
        file = "Analise_Revisoes_Yamaha.xlsx"
        df_cons = aplicar_filtros_globais(ler_dados_nps_oficial(file, "Consolidado_PBI_Revisoes"), f_causa, f_sub, f_nps, f_imp)
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
# FOOTER / INFO DA MARCA
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