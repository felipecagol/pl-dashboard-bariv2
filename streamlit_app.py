import re
import unicodedata
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="Dashboard P&L 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ARQUIVO_PADRAO = "BASE_DASHBOARD_PL_2026.xlsx"
ABA_RESULTADO = "RESULTADO"
ABA_BASE = "BASE_DASH"
DATA_MINIMA_DASH = pd.Timestamp(2026, 1, 1)

CSS = """
<style>
    .stApp { background: #080f1f; color: #e5ecff; }
    [data-testid="stSidebar"] { background: #0b1224; border-right: 1px solid #1e2a44; }
    [data-testid="stHeader"] { background: rgba(8, 15, 31, .95); }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
    .dash-title { font-size: 2.25rem; font-weight: 850; color: #ffffff; letter-spacing: .2px; margin-bottom: .2rem; }
    .dash-subtitle { color: #ffffff; font-size: .95rem; margin-bottom: 1.3rem; }
    .section-title { color: #ffffff; font-size: 1.25rem; font-weight: 750; margin-top: 1.1rem; margin-bottom: .6rem; }
    .kpi-card {
        background: #111a2e;
        border: 1px solid #243150;
        border-radius: 16px;
        padding: 22px 22px;
        min-height: 138px;
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
        text-align: center;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .card-row-spacer {
        height: 14px;
    }
    table.pnl-matrix {
        width: 100%;
        border-collapse: collapse;
        background: #080f1f;
        color: #e5ecff;
        font-size: .88rem;
    }
    table.pnl-matrix thead th {
        background: #111a2e;
        color: #ffffff;
        font-weight: 850;
        text-align: center;
        padding: 11px 10px;
        border-right: 1px solid rgba(255,255,255,.55);
        border-bottom: 1px solid rgba(255,255,255,.72);
        white-space: nowrap;
    }
    table.pnl-matrix tbody td {
        background: #080f1f;
        color: #e5ecff;
        text-align: center;
        vertical-align: middle;
        padding: 10px 10px;
        border-right: 1px solid rgba(255,255,255,.38);
        border-bottom: 1px solid rgba(255,255,255,.28);
        white-space: nowrap;
        font-weight: 400;
    }
    table.pnl-matrix tbody td:first-child {
        text-align: left;
        color: #ffffff;
        font-weight: 850;
        min-width: 200px;
    }
    table.pnl-matrix tbody tr.main-line td {
        background: #162338;
        color: #ffffff;
        font-weight: 850;
    }
    table.pnl-matrix tbody tr.main-line td:first-child {
        font-weight: 900;
    }
    table.pnl-matrix tbody tr.result-line td {
        background: #1d2d48;
        color: #ffffff;
        font-weight: 950;
        font-size: .95rem;
    }
    table.pnl-matrix td.neg-value {
        color: #ef4444;
        font-weight: 400;
    }
    table.pnl-matrix tbody tr.result-line td.neg-value {
        font-weight: 950;
    }
    table.pnl-matrix td.delta-positive,
    table.pnl-matrix tbody tr.main-line td.delta-positive {
        color: #22c55e !important;
        font-weight: 400 !important;
    }
    table.pnl-matrix td.delta-negative,
    table.pnl-matrix tbody tr.main-line td.delta-negative {
        color: #ef4444 !important;
        font-weight: 400 !important;
    }
    table.pnl-matrix tbody tr.result-line td.delta-positive {
        color: #22c55e !important;
        font-weight: 950 !important;
    }
    table.pnl-matrix tbody tr.result-line td.delta-negative {
        color: #ef4444 !important;
        font-weight: 950 !important;
    }
    table.pnl-matrix th.product-header {
        background: #101a2d;
        font-size: .95rem;
        letter-spacing: .2px;
    }
    table.pnl-matrix th.metric-header {
        background: #162338;
        font-size: .84rem;
    }
    .kpi-label { color: #ffffff; font-size: .96rem; font-weight: 700; margin-bottom: 12px; }
    .kpi-value { color: #ffffff; font-size: 1.90rem; font-weight: 850; line-height: 1.15; }
    .kpi-help { color: #ffffff; font-size: .85rem; font-weight: 700; margin-top: 10px; }
    .side-card {
        background: #111a2e;
        border: 1px solid #243150;
        border-radius: 16px;
        padding: 28px 24px;
        min-height: 480px;
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
        text-align: center;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        margin-top: 10px;
        position: relative;
    }
    .side-card-label {
        color: #ffffff;
        font-size: 1.05rem;
        font-weight: 700;
        position: absolute;
        top: 28px;
        left: 24px;
        right: 24px;
        text-align: center;
    }
    .side-card-value {
        color: #ffffff;
        font-size: 2.6rem;
        font-weight: 900;
        line-height: 1.1;
        margin-top: 0;
    }
    .side-card-delta {
        font-size: 1.45rem;
        font-weight: 900;
        margin-top: 16px;
    }
    .side-card-help {
        color: #ffffff;
        font-size: .95rem;
        font-weight: 700;
        margin-top: 16px;
        line-height: 1.35;
    }

    .composition-card {
        min-height: 480px;
        align-items: stretch;
        text-align: left;
        padding: 28px 22px;
        justify-content: center;
        position: relative;
    }
    .composition-title {
        color: #ffffff;
        font-size: 1.0rem;
        font-weight: 700;
        text-align: center;
        position: absolute;
        top: 28px;Itaú: O gráfico de linhas agora possui as datas do eixo X (jan/2026, fev/2026, etc.) em **branco e negrito**. O valor da Carteira de Crédito para 1Q26 foi ajustado para **1.995.139.307** nos cards do comparativo, como solicitado. O resto do dashboard permanece inalterado.

```python
import re
import unicodedata
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="Dashboard P&L 2026",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ARQUIVO_PADRAO = "BASE_DASHBOARD_PL_2026.xlsx"
ABA_RESULTADO = "RESULTADO"
ABA_BASE = "BASE_DASH"
DATA_MINIMA_DASH = pd.Timestamp(2026, 1, 1)

CSS = """
<style>
    .stApp { background: #080f1f; color: #e5ecff; }
    [data-testid="stSidebar"] { background: #0b1224; border-right: 1px solid #1e2a44; }
    [data-testid="stHeader"] { background: rgba(8, 15, 31, .95); }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
    .dash-title { font-size: 2.25rem; font-weight: 850; color: #ffffff; letter-spacing: .2px; margin-bottom: .2rem; }
    .dash-subtitle { color: #ffffff; font-size: .95rem; margin-bottom: 1.3rem; }
    .section-title { color: #ffffff; font-size: 1.25rem; font-weight: 750; margin-top: 1.1rem; margin-bottom: .6rem; }
    .kpi-card {
        background: #111a2e;
        border: 1px solid #243150;
        border-radius: 16px;
        padding: 22px 22px;
        min-height: 138px;
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
        text-align: center;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }
    .card-row-spacer {
        height: 14px;
    }
    table.pnl-matrix {
        width: 100%;
        border-collapse: collapse;
        background: #080f1f;
        color: #e5ecff;
        font-size: .88rem;
    }
    table.pnl-matrix thead th {
        background: #111a2e;
        color: #ffffff;
        font-weight: 850;
        text-align: center;
        padding: 11px 10px;
        border-right: 1px solid rgba(255,255,255,.55);
        border-bottom: 1px solid rgba(255,255,255,.72);
        white-space: nowrap;
    }
    table.pnl-matrix tbody td {
        background: #080f1f;
        color: #e5ecff;
        text-align: center;
        vertical-align: middle;
        padding: 10px 10px;
        border-right: 1px solid rgba(255,255,255,.38);
        border-bottom: 1px solid rgba(255,255,255,.28);
        white-space: nowrap;
        font-weight: 400;
    }
    table.pnl-matrix tbody td:first-child {
        text-align: left;
        color: #ffffff;
        font-weight: 850;
        min-width: 200px;
    }
    table.pnl-matrix tbody tr.main-line td {
        background: #162338;
        color: #ffffff;
        font-weight: 850;
    }
    table.pnl-matrix tbody tr.main-line td:first-child {
        font-weight: 900;
    }
    table.pnl-matrix tbody tr.result-line td {
        background: #1d2d48;
        color: #ffffff;
        font-weight: 950;
        font-size: .95rem;
    }
    table.pnl-matrix td.neg-value {
        color: #ef4444;
        font-weight: 400;
    }
    table.pnl-matrix tbody tr.result-line td.neg-value {
        font-weight: 950;
    }
    table.pnl-matrix td.delta-positive,
    table.pnl-matrix tbody tr.main-line td.delta-positive {
        color: #22c55e !important;
        font-weight: 400 !important;
    }
    table.pnl-matrix td.delta-negative,
    table.pnl-matrix tbody tr.main-line td.delta-negative {
        color: #ef4444 !important;
        font-weight: 400 !important;
    }
    table.pnl-matrix tbody tr.result-line td.delta-positive {
        color: #22c55e !important;
        font-weight: 950 !important;
    }
    table.pnl-matrix tbody tr.result-line td.delta-negative {
        color: #ef4444 !important;
        font-weight: 950 !important;
    }
    table.pnl-matrix th.product-header {
        background: #101a2d;
        font-size: .95rem;
        letter-spacing: .2px;
    }
    table.pnl-matrix th.metric-header {
        background: #162338;
        font-size: .84rem;
    }
    .kpi-label { color: #ffffff; font-size: .96rem; font-weight: 700; margin-bottom: 12px; }
    .kpi-value { color: #ffffff; font-size: 1.90rem; font-weight: 850; line-height: 1.15; }
    .kpi-help { color: #ffffff; font-size: .85rem; font-weight: 700; margin-top: 10px; }
    .side-card {
        background: #111a2e;
        border: 1px solid #243150;
        border-radius: 16px;
        padding: 28px 24px;
        min-height: 480px;
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
        text-align: center;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        margin-top: 10px;
        position: relative;
    }
    .side-card-label {
        color: #ffffff;
        font-size: 1.05rem;
        font-weight: 700;
        position: absolute;
        top: 28px;
        left: 24px;
        right: 24px;
        text-align: center;
    }
    .side-card-value {
        color: #ffffff;
        font-size: 2.6rem;
        font-weight: 900;
        line-height: 1.1;
        margin-top: 0;
    }
    .side-card-delta {
        font-size: 1.45rem;
        font-weight: 900;
        margin-top: 16px;
    }
    .side-card-help {
        color: #ffffff;
        font-size: .95rem;
        font-weight: 700;
        margin-top: 16px;
        line-height: 1.35;
    }

    .composition-card {
        min-height: 480px;
        align-items: stretch;
        text-align: left;
        padding: 28px 22px;
        justify-content: center;
        position: relative;
    }
    .composition-title {
        color: #ffffff;
        font-size: 1.0st.sidebar.title("Filtros")
arquivo_local = Path(ARQUIVO_PADRAO)
upload = st.sidebar.file_uploader("Atualizar base manualmente", type=["xlsx"])

if upload is not None:
    arquivo = upload
elif arquivo_local.exists():
    arquivo = arquivo_local
else:
    st.error(f"Arquivo '{ARQUIVO_PADRAO}' não encontrado no repositório.")
    st.stop()

try:
    df_resultado = carregar_resultado(arquivo)
except Exception as erro:
    st.error(f"Erro ao carregar a aba RESULTADO: {erro}")
    st.stop()

try:
    df_pnl_completo_global = carregar_pnl_mensal(arquivo)
    df_pnl_completo_global = garantir_linha_despesas_administrativas(df_pnl_completo_global)
    erro_pnl_global = None
except Exception as erro_pnl:
    df_pnl_completo_global = pd.DataFrame()
    erro_pnl_global = erro_pnl

periodos_disponiveis = (
    df_resultado[["Data", "Período"]]
    .drop_duplicates()
    .sort_values("Data")
    .reset_index(drop=True)
)

periodo_padrao = len(periodos_disponiveis) - 1

st.sidebar.markdown(
    """
    <div class="note-box">
        Exibindo somente dados a partir de <b>jan/2026</b>.
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="dash-title">Dashboard P&L 2026</div>', unsafe_allow_html=True)
st.markdown('<div class="dash-subtitle">Resultados consolidados, evolução histórica e abertura por empresa.</div>', unsafe_allow_html=True)

tab_resultados, tab_pnl_mensal, tab_pnl_acum, tab_comp_2025 = st.tabs(
    ["Resultados", "P&L Mensal", "P&L Acumulado", "Comparativo 2025"]
)

with tab_resultados:
    st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
    col_filtro_mes, col_filtro_vazio = st.columns([1, 3])
    with col_filtro_mes:
        periodo_sel = st.selectbox(
            "Mês de referência",
            periodos_disponiveis["Período"].tolist(),
            index=periodo_padrao,
            key="periodo_resultados",
        )

    empresa_sel_result = "Todos"

    df_principais = montar_resultados_principais(df_resultado)
    df_cards = df_principais[df_principais["Período"] == periodo_sel].copy()
    periodo_ant = periodo_anterior(periodos_disponiveis, periodo_sel)

    st.markdown('<div class="section-title">Principais resultados</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    indicadores = [
        "Resultado Conglomerado Financeiro",
        "Resultado Coligadas",
        "Resultado Conglomerado + Coligadas",
        "Resultado Total",
    ]

    for coluna, indicador in zip([c1, c2, c3, c4], indicadores):
        with coluna:
            linha = df_cards[df_cards["Indicador"] == indicador]
            valor = linha["Valor"].sum() if not linha.empty else 0
            variacao = variacao_mes_anterior(df_principais, indicador, periodo_sel, periodo_ant)
            card(indicador, valor, variacao=variacao)

    st.markdown('<div class="section-title">Evolução histórica dos resultados</div>', unsafe_allow_html=True)

    if df_principais.empty:
        st.warning("Não encontrei as linhas principais na aba RESULTADO. Verifique os nomes das linhas na planilha.")
    else:
        base_linhas = df_principais.sort_values(["Indicador", "Data"]).copy()
        base_linhas["Rótulo"] = base_linhas["Valor"].map(formatar_moeda_curta)

        fig = px.line(
            base_linhas,
            x="Data",
            y="Valor",
            color="Indicador",
            markers=True,
            line_shape="spline",
            labels={"Data": "Mês", "Valor": "Resultado", "Indicador": "Resultado"},
        )

        for trace in fig.data:
            trace.update(
                mode="lines+markers",
                cliponaxis=False,
            )

        offsets_rotulo = {
            "Resultado Total": {"xshift": 0, "yshift": 18, "xanchor": "center"},
            "Resultado Conglomerado + Coligadas": {"xshift": 0, "yshift": -18, "xanchor": "center"},
            "Resultado Conglomerado Financeiro": {"xshift": 0, "yshift": 18, "xanchor": "center"},
            "Resultado Coligadas": {"xshift": 0, "yshift": -18, "xanchor": "center"},
        }

        for _, row in base_linhas.iterrows():
            desloc = offsets_rotulo.get(row["Indicador"], {"xshift": 0, "yshift": 18, "xanchor": "center"})
            fig.add_annotation(
                x=row["Data"],
                y=row["Valor"],
                text=f"<b>{row['Rótulo']}</b>",
                showarrow=False,
                xshift=desloc["xshift"],
                yshift=desloc["yshift"],
                font=dict(size=12, color="#FFFFFF", family="Arial Black"),
                xanchor=desloc["xanchor"],
                align="center",
                bgcolor="rgba(0,0,0,0)",
            )

        tick_datas = periodos_disponiveis["Data"].tolist()
        tick_textos = periodos_disponiveis["Período"].tolist()

        y_min = base_linhas["Valor"].min()
        y_max = base_linhas["Valor"].max()
        y_pad = max((y_max - y_min) * 0.24, 1)

        x_min = min(tick_datas) - pd.DateOffset(days=8)
        x_max = max(tick_datas) + pd.DateOffset(days=20)

        fig.update_layout(
            template="plotly_dark",
            paper_bgcolor="#080f1f",
            plot_bgcolor="#080f1f",
            height=500,
            margin=dict(l=10, r=40, t=35, b=20),
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="left", x=0, font=dict(size=13, color="#ffffff", family="Arial Black")),
        )
        fig.update_xaxes(
            tickmode="array",
            tickvals=tick_datas,
            ticktext=tick_textos,
            range=[x_min, x_max],
            title_text="",
            showgrid=False,
            zeroline=False,
            tickfont=dict(color='#FFFFFF', size=13, family="Arial Black"),
        )
        fig.update_yaxes(
            tickprefix="R$ ",
            separatethousands=True,
            range=[y_min - y_pad, y_max + y_pad],
            title_text="",
            showgrid=False,
            zeroline=False,
        )
        st.plotly_chart(fig, use_container_width=True)

        col_acum, col_comp, col_gauge = st.columns([1, 1, 1.6])

        valor_acumulado, variacao_acumulado, valor_acumulado_anterior, data_inicio = resultado_total_acumulado_ano(
            df_principais, periodo_sel
        )

        with col_acum:
            card_resultado_total_acumulado(valor_acumulado, variacao_acumulado, valor_acumulado_anterior, periodo_sel)

        with col_comp:
            card_composicao_resultado_total_acumulado(df_pnl_completo_global, periodo_sel, empresa_sel_result)

        with col_gauge:
            ORCADO_2026 = 73_400_000.0
            fig_gauge = grafico_alcance_vs_orcado(valor_acumulado, ORCADO_2026)
            if fig_gauge is not None:
                st.plotly_chart(fig_gauge, use_container_width=True)

    st.markdown('<div class="section-title">Resultado aberto por empresa</div>', unsafe_allow_html=True)

    tabela = montar_tabela_empresas_e_total(df_resultado)
    tabela = filtrar_tabela_resultado_por_empresa(tabela, empresa_sel_result)
    tabela, coluna_delta = adicionar_coluna_variacao_tabela(tabela, periodos_disponiveis, periodo_sel)

    colunas_periodo = [c for c in tabela.columns if c not in ["Linha", coluna_delta]]
    tabela["Acumulado"] = tabela[colunas_periodo].apply(
        lambda row: pd.to_numeric(row, errors="coerce").sum(), axis=1
    )

    outras_cols = [c for c in tabela.columns if c not in ["Linha", "Acumulado", coluna_delta]]
    tabela = tabela[["Linha"] + outras_cols + ["Acumulado", coluna_delta]]
    is_total = tabela["Linha"].str.lower().str.contains("resultado total", na=False)
    tabela_corpo = tabela[~is_total].sort_values("Acumulado", ascending=False)
    tabela_total = tabela[is_total]
    tabela = pd.concat([tabela_corpo, tabela_total], ignore_index=True)

    tabela_valores = tabela.copy()
    tabela_formatada = tabela.copy()

    for col in tabela_formatada.columns:
        if col == coluna_delta:
            tabela_formatada[col] = tabela_formatada[col].map(formatar_percentual)
        elif col == "Acumulado":
            tabela_formatada[col] = tabela_formatada[col].map(formatar_numero)
        elif col != "Linha":
            tabela_formatada[col] = tabela_formatada[col].map(formatar_numero)

    tabela_formatada = tabela_formatada.rename(columns={"Linha": "Empresa"})
    tabela_valores = tabela_valores.rename(columns={"Linha": "Empresa"})

    st.markdown(tabela_html(tabela_formatada, tabela_valores, coluna_delta=coluna_delta), unsafe_allow_html=True)

with tab_pnl_mensal:
    if erro_pnl_global is None and not df_pnl_completo_global.empty:
        render_pnl_page(df_pnl_completo_global, arquivo, pagina="Mensal")
    else:
        st.info(f"Não consegui carregar a aba P&L Mensal: {erro_pnl_global}")

with tab_pnl_acum:
    if erro_pnl_global is None and not df_pnl_completo_global.empty:
        try:
            df_comp_para_acum = carregar_comparativo_2025(arquivo)
        except Exception:
            df_comp_para_acum = None
        render_pnl_page(df_pnl_completo_global, arquivo, pagina="Acumulado", df_comp_2025=df_comp_para_acum)
    else:
        st.info(f"Não consegui carregar a aba P&L Acumulado: {erro_pnl_global}")

with tab_comp_2025:
    try:
        df_comp = carregar_comparativo_2025(arquivo)
        
        # Correção do valor da Carteira de Crédito para 1Q26 solicitada pelo usuário
        if df_comp is not None and not df_comp.empty:
            # Encontrar todas as linhas de carteira para 2026, Total e substituir o valor
            df_comp.loc[
                (df_comp["Ano"] == 2026) & 
                (df_comp["Produto"] == "Total") & 
                (df_comp["Linha_Normalizada"].str.contains("carteira", na=False)), 
                "Realizado"
            ] = 1995139307.0
            
        df_2025_acumulado = carregar_2025_acumulado(arquivo)
        df_comp_principais = montar_comparativo_principais(df_comp, df_2025_acumulado)

        if df_comp.empty or df_comp_principais.empty:
            st.info("Não encontrei dados suficientes na aba Comparativo 2026 x 2025.")
        else:
            st.markdown('<div class="section-title">Comparativo 1Q26 x 1Q25</div>', unsafe_allow_html=True)
            st.markdown('<div style="color: #ffffff; font-weight: bold; font-size: 1.1rem; margin-bottom: 20px; margin-top: -5px;">Legenda: Q = Quadrimestre</div>', unsafe_allow_html=True)

            novos_cards_linha1 = [
                ("Margem de Intermediação 1Q26", "MARGEM INTERMEDIAÇÃO"),
                ("MG Intermediação Líq. 1Q26", "MG INTERMEDIAÇÃO LIQ"),
                ("MG Contribuição Direta 1Q26", "MG CONTRIBUIÇÃO DIRETA"),
                ("Resultado Antes do Imposto 1Q26", "RESULTADO ANTES IMPOSTO"),
            ]
            
            cols1 = st.columns(4)
            for col, (titulo, linha_nome) in zip(cols1, novos_cards_linha1):
                with col:
                    linha_df = obter_linha_comparativo(df_comp_principais, linha_nome)
                    if linha_df.empty:
                        card(titulo, 0, ajuda="", variacao=None)
                    else:
                        valor_2025 = linha_df["2025"].iloc[0]
                        valor_2026 = linha_df["2026"].iloc[0]
                        variacao = linha_df["Δ %"].iloc[0]
                        ajuda = ""
                        cor_classe = None
                        variacao_exibir = None
                        if variacao is not None and not pd.isna(variacao) and pd.notna(valor_2026) and float(valor_2026) < 0 and float(variacao) < 0:
                            variacao_exibir = -variacao
                            cor_classe = "delta-negative"
                        card(titulo, valor_2026, ajuda=ajuda, variacao=variacao,
                             variacao_label="vs 1Q25", cor_classe=cor_classe, variacao_exibir=variacao_exibir)

            st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)

            novos_cards_linha2 = [
                ("Resultado Contábil 1Q26", "RESULTADO CONTÁBIL"),
                ("Carteira de Crédito (Média) 1Q26", "Carteira de Crédito Média"),
                ("PL Médio 1Q26", "PL Médio"),
            ]
            
            cols2 = st.columns(3)
            for col, (titulo, linha_nome) in zip(cols2, novos_cards_linha2):
                with col:
                    linha_df = obter_linha_comparativo(df_comp_principais, linha_nome)
                        
                    if linha_df.empty:
                        card(titulo, 0, ajuda="", variacao=None)
                    else:
                        valor_2025 = linha_df["2025"].iloc[0]
                        valor_2026 = linha_df["2026"].iloc[0]
                        variacao = linha_df["Δ %"].iloc[0]
                        ajuda = ""
                        cor_classe = None
                        variacao_exibir = None
                        if variacao is not None and not pd.isna(variacao) and pd.notna(valor_2026) and float(valor_2026) < 0 and float(variacao) < 0:
                            variacao_exibir = -variacao
                            cor_classe = "delta-negative"
                        card(titulo, valor_2026, ajuda=ajuda, variacao=variacao,
                             variacao_label="vs 1Q25", cor_classe=cor_classe, variacao_exibir=variacao_exibir)

            st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)

            st.markdown('<div class="section-title">Quanto do Resultado Contábil acumulado de 2025 já foi alcançado</div>', unsafe_allow_html=True)
            linha_resultado = obter_linha_comparativo(df_comp_principais, "RESULTADO CONTÁBIL")
            if linha_resultado.empty:
                st.info("Não encontrei a linha de Resultado Contábil para montar o gráfico de alcance.")
            else:
                valor_base_2025 = linha_resultado["2025 Acumulado"].iloc[0]
                valor_1t26 = linha_resultado["2026"].iloc[0]
                fig_alcance_resultado = grafico_alcance_resultado_contabil(valor_1t26, valor_base_2025)
                if fig_alcance_resultado is not None:
                    st.plotly_chart(fig_alcance_resultado, use_container_width=True)
                else:
                    st.info("Não foi possível calcular o alcance do Resultado Contábil com a base atual.")

            st.markdown('<div class="section-title">1Q25 x 1Q26 por linha principal</div>', unsafe_allow_html=True)
            
            linhas_ocultas_grafico = ["carteira de credito media", "pl medio", "carteira de credito bruta media", "pl medio banco hipo"]
            df_comp_grafico = df_comp_principais[~df_comp_principais["Linha"].map(normalizar_texto).isin(linhas_ocultas_grafico)]

            base_long = df_comp_grafico.melt(
                id_vars=["Linha", "Ordem"],
                value_vars=["2025", "2026"],
                var_name="Ano",
                value_name="Valor",
            ).dropna(subset=["Valor"])
            base_long["Rótulo"] = base_long["Valor"].map(formatar_moeda_curta)
            base_long = base_long.sort_values("Ordem", ascending=False)

            fig_comp_ano = px.bar(
                base_long,
                x="Valor",
                y="Linha",
                color="Ano",
                text="Rótulo",
                orientation="h",
                barmode="group",
                labels={"Valor": "", "Linha": "", "Ano": ""},
            )
            fig_comp_ano.update_traces(
                texttemplate="<b>%{text}</b>",
                textposition="outside",
                textfont=dict(size=11, color="#FFFFFF", family="Arial Black"),
                cliponaxis=False,
            )
            if not base_long.empty:
                xmin = base_long["Valor"].min()
                xmax = base_long["Valor"].max()
                xpad = max((xmax - xmin) * 0.18, 1)
                fig_comp_ano.update_xaxes(range=[xmin - xpad, xmax + xpad], showgrid=False, zeroline=False, tickprefix="R$ ", separatethousands=True)
            fig_comp_ano.update_yaxes(showgrid=False, zeroline=False)
            fig_comp_ano.update_layout(
                template="plotly_dark",
                paper_bgcolor="#080f1f",
                plot_bgcolor="#080f1f",
                height=520,
                margin=dict(l=10, r=120, t=40, b=20),
                title={"text": "<b>1Q25 x 1Q26 por linha principal</b>", "font": {"size": 16, "color": "#ffffff"}},
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=13, color="#ffffff", family="Arial Black")),
            )
            st.plotly_chart(fig_comp_ano, use_container_width=True)

            st.markdown('<div class="section-title">Tabela comparativa</div>', unsafe_allow_html=True)
            st.markdown(tabela_html_comparativo(df_comp_principais), unsafe_allow_html=True)

    except Exception as erro:
        st.info(f"Não consegui carregar a aba Comparativo 2026 x 2025: {erro}")
