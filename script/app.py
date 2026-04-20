# ==========================================
# FUNÇÃO DE CORREÇÃO DO GRÁFICO DE IMPACTO
# ==========================================
def gerar_grafico_impacto_corrigido(df_plot, col_y):
    """
    Solução Definitiva: Converte os dados do Pandas para Listas Nativas do Python.
    Isso impede qualquer bug de renderização de eixos ou truncamento numérico no GCP.
    """
    if df_plot.empty or 'Impacto' not in df_plot.columns:
        fig_vazia = go.Figure()
        fig_vazia.update_layout(title="Sem dados para exibir", plot_bgcolor='rgba(0,0,0,0)')
        return fig_vazia

    # 1. Cópia segura e garantia numérica
    df_plot = df_plot.copy()
    df_plot['Impacto'] = pd.to_numeric(df_plot['Impacto'], errors='coerce').fillna(0.0)
    df_plot = df_plot.sort_values('Impacto', ascending=True)

    # 2. O PULO DO GATO: Converter para Listas Python puras!
    # Isso impede o bug do eixo Y (0 a 18) e o truncamento do eixo X
    eixo_y = df_plot[col_y].astype(str).tolist()
    valores_x = df_plot['Impacto'].tolist()

    # 3. Montagem Manual de Cores e Textos (para não depender da inteligência da biblioteca)
    cores = []
    textos = []
    for val in valores_x:
        # Cor: Laranja da Insights&Etc para negativo, Turquesa para positivo
        cores.append("#FF6B6B" if val < 0 else "#00D2D3")
        
        # Texto: Como os valores (ex: -0.015) são pequenos, mostramos 3 casas decimais
        if abs(val) > 0 and abs(val) < 0.1:
            textos.append(f"{val:.3f}")
        else:
            textos.append(f"{val:.1f}")

    # 4. Construção Blindada via Graph Objects
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=valores_x,
        y=eixo_y,
        orientation='h',
        marker_color=cores,
        text=textos,
        textposition='outside',
        cliponaxis=False
    ))

    # 5. Eixo X 100% Simétrico
    max_abs = max([abs(v) for v in valores_x]) if valores_x else 1.0
    max_range = (max_abs * 1.3) if max_abs > 0 else 1.0

    fig.update_layout(
        xaxis=dict(
            range=[-max_range, max_range], 
            zeroline=True, 
            zerolinewidth=2, 
            zerolinecolor='rgba(0,0,0,0.3)',
            showgrid=False
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=10, r=50) # Margem esquerda pequena, pois os nomes vão naturalmente para a esquerda
    )

    return fig