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
    .kpi-delta {
        font-size: .98rem;
        font-weight: 800;
        margin-top: 10px;
        line-height: 1.15;
    }
    .delta-positive { color: #22c55e; }
    .delta-negative { color: #ef4444; }
    .delta-neutral { color: #ffffff; }
    .note-box {
        background: #111a2e;
        border: 1px solid #243150;
        border-radius: 14px;
        padding: 13px 16px;
        color: #ffffff;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; border-bottom: 1px solid #243150; }
    .stTabs [data-baseweb="tab"] { color: #ffffff; background: transparent; }
    .stTabs [aria-selected="true"] { color: #ffffff; border-bottom: 2px solid #24a8ff; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def formatar_moeda(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    sinal = "-" if valor < 0 else ""
    valor_abs = abs(valor)
    if valor_abs >= 1_000_000_000:
        texto = f"{sinal}R$ {valor_abs / 1_000_000_000:,.2f} bi"
    elif valor_abs >= 1_000_000:
        texto = f"{sinal}R$ {valor_abs / 1_000_000:,.2f} mi"
    elif valor_abs >= 1_000:
        texto = f"{sinal}R$ {valor_abs / 1_000:,.1f} mil"
    else:
        texto = f"{sinal}R$ {valor_abs:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_moeda_curta(valor):
    try:
        valor = float(valor)
    except Exception:
        return ""
    sinal = "-" if valor < 0 else ""
    valor_abs = abs(valor)
    if valor_abs >= 1_000_000_000:
        texto = f"{sinal}R$ {valor_abs / 1_000_000_000:,.1f} bi"
    elif valor_abs >= 1_000_000:
        texto = f"{sinal}R$ {valor_abs / 1_000_000:,.1f} mi"
    elif valor_abs >= 1_000:
        texto = f"{sinal}R$ {valor_abs / 1_000:,.0f} mil"
    else:
        texto = f"{sinal}R$ {valor_abs:,.0f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_numero(valor):
    if pd.isna(valor): return ""
    try:
        valor = float(valor)
    except Exception: return str(valor)
    return f"{valor:,.0f}".replace(",", ".")


def formatar_percentual(valor):
    if pd.isna(valor): return ""
    try:
        valor = float(valor)
    except Exception: return str(valor)
    sinal = "+" if valor > 0 else ""
    texto = f"{sinal}{valor * 100:,.1f}%"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def linhas_percentuais_pnl():
    return {
        normalizar_texto("Margem Bruta"),
        normalizar_texto("Margem Liquida"),
        normalizar_texto("Margem Líquida"),
        normalizar_texto("Rácio de Eficiência"),
        normalizar_texto("Rácio de Eficiência Recorrente"),
        normalizar_texto("RPL - RES. CONTÁBIL"),
        normalizar_texto("Taxa Média Carteira Bruta Média"),
        normalizar_texto("Taxa Média Carteira SD Cliente Média"),
        normalizar_texto("Rateio Carteira"),
        normalizar_texto("Rateio da Carteira"),
        normalizar_texto("Alíquota de IR/CSLL"),
    }


def linha_pnl_percentual(linha):
    return normalizar_texto(linha) in linhas_percentuais_pnl()


def formatar_percentual_valor(valor):
    if pd.isna(valor): return ""
    try:
        valor = float(valor)
    except Exception: return str(valor)
    texto = f"{valor * 100:,.1f}%"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_pontos_percentuais(valor):
    if pd.isna(valor): return ""
    try:
        valor = float(valor)
    except Exception: return str(valor)
    sinal = "+" if valor > 0 else ""
    texto = f"{sinal}{valor * 100:,.1f} pp"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def converter_periodo(valor):
    if pd.isna(valor): return None
    if isinstance(valor, pd.Timestamp): return valor.to_period("M").to_timestamp()
    texto = str(valor).strip().lower()
    if not texto or texto == "nan": return None
    meses = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6, "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
    texto_sem_acento = normalizar_texto(texto)
    partes = texto_sem_acento.split()
    mes, ano = None, None
    for parte in partes:
        if parte in meses: mes = meses[parte]
        elif re.fullmatch(r"\d{4}", parte): ano = int(parte)
        elif re.fullmatch(r"\d{2}", parte): ano = 2000 + int(parte)
    if mes and ano: return pd.Timestamp(ano, mes, 1)
    return pd.to_datetime(texto, errors="coerce", dayfirst=True)


def nome_periodo(data):
    if pd.isna(data): return ""
    meses = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    data = pd.Timestamp(data)
    return f"{meses[data.month - 1]}/{data.year}"


def formatar_variacao(valor, label="Δ mês anterior"):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    sinal = "+" if valor > 0 else ""
    texto = f"{label} {sinal}{valor * 100:,.1f}%"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def classe_variacao(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    if valor > 0: return "delta-positive"
    if valor < 0: return "delta-negative"
    return "delta-neutral"


def card(titulo, valor, ajuda="", variacao=None, variacao_label="Δ mês anterior", cor_classe=None, variacao_exibir=None):
    delta_html = ""
    if variacao is not None:
        cls = cor_classe if cor_classe is not None else classe_variacao(variacao)
        val_txt = variacao_exibir if variacao_exibir is not None else variacao
        delta_html = f'<div class="kpi-delta {cls}">{formatar_variacao(val_txt, variacao_label)}</div>'
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{titulo}</div>
            <div class="kpi-value">{formatar_moeda(valor)}</div>
            {delta_html}
            {f'<div class="kpi-help">{ajuda}</div>' if ajuda else ""}
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def carregar_resultado(arquivo):
    bruto = pd.read_excel(arquivo, sheet_name=ABA_RESULTADO, header=None, engine="openpyxl")
    linha_mes, col_rotulo = None, None
    for idx in bruto.index:
        for col in bruto.columns:
            if normalizar_texto(bruto.loc[idx, col]) == "mes":
                linha_mes, col_rotulo = idx, col
                break
        if linha_mes is not None: break
    colunas_periodo = []
    for col in bruto.columns:
        if col <= col_rotulo: continue
        periodo = converter_periodo(bruto.loc[linha_mes, col])
        if periodo is not None: colunas_periodo.append((col, periodo))
    registros = []
    ordem_linha = 0
    for idx in bruto.index:
        if idx <= linha_mes: continue
        linha_nome = bruto.loc[idx, col_rotulo]
        if pd.isna(linha_nome) or str(linha_nome).strip() == "": continue
        for col, periodo in colunas_periodo:
            valor = pd.to_numeric(bruto.loc[idx, col], errors="coerce")
            if pd.notna(valor):
                registros.append({"Linha": str(linha_nome).strip(), "Linha_Normalizada": normalizar_texto(linha_nome), "Data": periodo, "Período": nome_periodo(periodo), "Valor": float(valor), "Ordem_Linha": ordem_linha})
        ordem_linha += 1
    df = pd.DataFrame(registros)
    return df[df["Data"] >= DATA_MINIMA_DASH].copy()


def obter_periodos_pnl_mensal_anualizado(arquivo):
    try: bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal - Anualizado", header=None, engine="openpyxl")
    except Exception: bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal", header=None, engine="openpyxl")
    periodos = []
    chaves_vistas = set()
    for idx in bruto.index:
        for col in bruto.columns:
            if normalizar_texto(bruto.loc[idx, col]) != "data base": continue
            for c_data in range(col + 1, min(col + 12, max(bruto.columns) + 1)):
                if c_data not in bruto.columns: continue
                data = converter_periodo(bruto.loc[idx, c_data])
                if data is None: continue
                data_ts = pd.Timestamp(data)
                chave = data_ts.strftime("%Y-%m")
                if chave not in chaves_vistas:
                    periodos.append({"Período": nome_periodo(data_ts), "Data": data_ts})
                    chaves_vistas.add(chave)
                break
    return sorted(periodos, key=lambda x: x["Data"])


@st.cache_data(show_spinner=False)
def carregar_pnl_mensal(arquivo):
    try: bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal - Anualizado", header=None, engine="openpyxl")
    except Exception: bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal", header=None, engine="openpyxl")
    registros = []
    for idx in bruto.index:
        for col in bruto.columns:
            if normalizar_texto(bruto.loc[idx, col]) != "data base": continue
            data_base = None
            for c_data in range(col + 1, min(col + 12, max(bruto.columns) + 1)):
                if c_data in bruto.columns:
                    data_base = converter_periodo(bruto.loc[idx, c_data])
                    if data_base is not None: break
            if data_base is None: continue
            blocos = []
            for c_prod in range(col + 1, min(col + 12, max(bruto.columns) + 1)):
                prod_norm = normalizar_texto(bruto.loc[idx + 2, c_prod])
                if prod_norm in ["consignado", "imobiliario", "total"]:
                    produto = {"consignado": "Consignado", "imobiliario": "Imobiliário", "total": "Total"}[prod_norm]
                    for c_met in range(c_prod, min(c_prod + 4, max(bruto.columns) + 1)):
                        met_norm = normalizar_texto(bruto.loc[idx + 3, c_met])
                        if met_norm == "realizado": metrica = "Realizado"
                        elif met_norm == "orcado": metrica = "Orçado"
                        elif "r" in met_norm and ("r" == met_norm or "rs" in met_norm): metrica = "Δ R$"
                        elif "%" in str(bruto.loc[idx + 3, c_met]) or "delta" in met_norm: metrica = "Δ %"
                        else: continue
                        blocos.append({"Produto": produto, "Coluna": c_met, "Métrica": metrica})
            ordem = 0
            for r in bruto.index:
                if r <= idx + 3: continue
                linha_nome = bruto.loc[r, col]
                if pd.isna(linha_nome) or str(linha_nome).strip() == "": continue
                for bloco in blocos:
                    valor = pd.to_numeric(bruto.loc[r, bloco["Coluna"]], errors="coerce")
                    if pd.notna(valor):
                        registros.append({"Periodo": nome_periodo(data_base), "Data": pd.Timestamp(data_base), "Produto": bloco["Produto"], "Linha": str(linha_nome).strip(), "Linha_Normalizada": normalizar_texto(linha_nome), "Métrica": bloco["Métrica"], "Valor": float(valor), "Ordem_Linha": ordem})
                ordem += 1
    df = pd.DataFrame(registros)
    df = df.drop_duplicates(subset=["Periodo", "Produto", "Linha_Normalizada", "Métrica"], keep="first").reset_index(drop=True)
    mask_modulo = df["Linha_Normalizada"].isin([normalizar_texto("Alíquota de IR/CSLL"), normalizar_texto("Rácio de Eficiência")]) & df["Métrica"].isin(["Realizado", "Orçado"])
    df.loc[mask_modulo, "Valor"] = df.loc[mask_modulo, "Valor"].abs()
    return df


@st.cache_data(show_spinner=False)
def carregar_pnl_acumulado_oficial_completo(arquivo):
    try: bruto = pd.read_excel(arquivo, sheet_name="P&L Acumulado", header=None, engine="openpyxl")
    except Exception: return pd.DataFrame()
    data_ate = pd.Timestamp(bruto.iloc[0, 8])
    periodo_label = nome_periodo(data_ate)
    cols_produto = {"Consignado": {"Realizado": 5, "Orçado": 6}, "Imobiliário": {"Realizado": 8, "Orçado": 9}, "Total": {"Realizado": 11, "Orçado": 12}}
    registros = []
    ordem = 0
    for r in bruto.index:
        if r <= 3: continue
        linha_nome = bruto.iloc[r, 1]
        if pd.isna(linha_nome) or str(linha_nome).strip() == "": continue
        linha_norm = normalizar_texto(linha_nome)
        for produto, cols in cols_produto.items():
            val_real = pd.to_numeric(bruto.iloc[r, cols["Realizado"]], errors="coerce")
            val_orc = pd.to_numeric(bruto.iloc[r, cols["Orçado"]], errors="coerce")
            if pd.notna(val_real): registros.append({"Periodo": periodo_label, "Data": data_ate, "Produto": produto, "Linha": str(linha_nome).strip(), "Linha_Normalizada": linha_norm, "Métrica": "Realizado", "Valor": float(val_real), "Ordem_Linha": ordem})
            if pd.notna(val_orc): registros.append({"Periodo": periodo_label, "Data": data_ate, "Produto": produto, "Linha": str(linha_nome).strip(), "Linha_Normalizada": linha_norm, "Métrica": "Orçado", "Valor": float(val_orc), "Ordem_Linha": ordem})
        ordem += 1
    df = pd.DataFrame(registros)
    df = df.drop_duplicates(subset=["Periodo", "Produto", "Linha_Normalizada", "Métrica"], keep="first").reset_index(drop=True)
    mask_modulo = df["Linha_Normalizada"].isin([normalizar_texto("Alíquota de IR/CSLL"), normalizar_texto("Rácio de Eficiência")]) & df["Métrica"].isin(["Realizado", "Orçado"])
    df.loc[mask_modulo, "Valor"] = df.loc[mask_modulo, "Valor"].abs()
    return df


def obter_linhas_principais_pnl(df_pnl):
    # === ADICIONADA NOVA LINHA DE CARDS NO TOPO ===
    linhas_desejadas = [
        "Impostos Diretos",
        "Comissão (Apoio - Comercial - Mídia Paga)",
        "Amortização ",
        "RECEITAS",
        "DESPESAS DE ORIGINAÇÃO",
        "MARGEM INTERMEDIAÇÃO",
        "Provisões",
        "MG INTERMEDIAÇÃO LIQ",
        "Despesas Administrativas",
        "RESULTADO ANTES IMPOSTO",
        "Impostos (IR/CSLL)",
        "RESULTADO CONTÁBIL",
    ]
    disponiveis = df_pnl[["Linha", "Linha_Normalizada", "Ordem_Linha"]].drop_duplicates().sort_values("Ordem_Linha")
    selecionadas = []
    for linha in linhas_desejadas:
        alvo = normalizar_texto(linha)
        match = disponiveis[disponiveis["Linha_Normalizada"].eq(alvo)]
        if match.empty: match = disponiveis[disponiveis["Linha_Normalizada"].str.contains(alvo, regex=False, na=False)]
        if not match.empty: selecionadas.append(match.iloc[0]["Linha"])
    return selecionadas


def render_pnl_page(df_pnl_completo, arquivo, pagina="Mensal", df_comp_2025=None):
    periodos_pnl = obter_periodos_pnl_mensal_anualizado(arquivo)
    if pagina == "Acumulado":
        meses_fim_trimestre = {3, 6, 9, 12}
        periodos_pnl = [p for p in periodos_pnl if p["Data"] is not None and p["Data"].month in meses_fim_trimestre]
        label_para_periodo = {f"{ {3:'1Q', 6:'2Q', 9:'3Q', 12:'4Q'}.get(p['Data'].month, p['Período']) } - {p['Data'].year}": p["Período"] for p in periodos_pnl if p["Data"] is not None}
        lista_periodos_pnl = list(label_para_periodo.keys())
        df_pnl = carregar_pnl_acumulado_oficial_completo(arquivo)
        df_pnl = garantir_linha_despesas_administrativas(df_pnl)
    else:
        lista_periodos_pnl = [item["Período"] for item in periodos_pnl]
        label_sel = st.sidebar.selectbox("Data base", lista_periodos_pnl, index=len(lista_periodos_pnl)-1, key=f"data_pnl_{pagina.lower()}")
        df_pnl = df_pnl_completo[df_pnl_completo["Periodo"] == label_sel].copy()
        data_sel_pnl = label_sel

    produto_sel_pnl = st.sidebar.selectbox("Produto", ["Consignado", "Imobiliário", "Total"], index=2, key=f"prod_pnl_{pagina.lower()}")
    linhas_principais = obter_linhas_principais_pnl(df_pnl)
    st.markdown(f'<div class="section-title">Principais linhas do P&L {pagina}</div>', unsafe_allow_html=True)
    
    for inicio in range(0, len(linhas_principais), 3):
        if inicio > 0: st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)
        cols = st.columns(3)
        for col, linha in zip(cols, linhas_principais[inicio:inicio+3]):
            realizado = valor_pnl(df_pnl, produto_sel_pnl, linha, "Realizado")
            if pagina == "Acumulado":
                variacao = variacao_pnl_acumulado_vs_2025(df_comp_2025, produto_sel_pnl, linha, realizado)
                v_label = "vs 1Q25"
            else:
                variacao = variacao_pnl_mes_anterior(df_pnl_completo, produto_sel_pnl, linha, data_sel_pnl)
                v_label = "vs mês ant."
            cor_classe, v_exibir = None, None
            if variacao is not None and not pd.isna(variacao) and realizado < 0:
                v_exibir, cor_classe = -variacao, ("delta-negative" if variacao < 0 else "delta-positive")
            with col: card(linha.replace("*","").strip(), realizado, variacao=variacao, variacao_label=v_label, cor_classe=cor_classe, variacao_exibir=v_exibir)

    st.markdown('<div class="section-title">Detalhamento</div>', unsafe_allow_html=True)
    linhas_tabela = obter_linhas_tabela_pnl(df_pnl)
    matriz, prods, mets = montar_matriz_pnl_excel(df_pnl, linhas_tabela)
    st.markdown(tabela_html_pnl_matriz(matriz, prods, mets), unsafe_allow_html=True)


# Funções de apoio (valor_pnl, montar_matriz_pnl_excel, etc continuam as mesmas das versões funcionais anteriores)
def valor_pnl(df, produto, linha, metrica):
    base = df[(df["Produto"] == produto) & (df["Linha"] == linha) & (df["Métrica"] == metrica)]
    return base.sort_values("Ordem_Linha")["Valor"].iloc[0] if not base.empty else 0

def montar_matriz_pnl_excel(df_pnl, linhas_principais):
    produtos = ["Consignado", "Imobiliário", "Total"]
    metricas_por_produto = {"Consignado": ["Realizado", "Orçado", "Δ %"], "Imobiliário": ["Realizado", "Orçado", "Δ %"], "Total": ["Realizado", "Orçado", "Δ %", "Δ R$"]}
    linhas = []
    for linha in linhas_principais:
        row = {"Linha": linha}
        is_pct, is_racio = linha_pnl_percentual(linha), normalizar_texto(linha) == normalizar_texto("Rácio de Eficiência")
        for produto in produtos:
            real, orc = valor_pnl(df_pnl, produto, linha, "Realizado"), valor_pnl(df_pnl, produto, linha, "Orçado")
            delta_rs = (orc - real) if is_racio else (real - orc)
            delta_pct = delta_rs if is_pct else (delta_rs / abs(orc) if orc != 0 else pd.NA)
            row[(produto, "Realizado")], row[(produto, "Orçado")], row[(produto, "Δ %")] = real, orc, delta_pct
            row[(produto, "_ambos_neg")] = (not is_racio and not is_pct and real < 0 and orc < 0)
            row[(produto, "_delta_bad")] = (False if pd.isna(delta_pct) or is_racio else ((-delta_pct if row[(produto, "_ambos_neg")] else delta_pct) < 0))
            if produto == "Total": row[(produto, "Δ R$")] = pd.NA if is_pct else delta_rs
        linhas.append(row)
    return pd.DataFrame(linhas), produtos, metricas_por_produto

def tabela_html_pnl_matriz(df_matrix, produtos, metricas_por_produto):
    html = ['<div class="table-wrap"><table class="pnl-matrix"><thead><tr><th rowspan="2">Linha P&L</th>']
    for p in produtos: html.append(f'<th class="product-header" colspan="{len(metricas_por_produto[p])}">{p.upper()}</th>')
    html.append('</tr><tr>')
    for p in produtos: 
        for m in metricas_por_produto[p]: html.append(f'<th class="metric-header">{m}</th>')
    html.append('</tr></thead><tbody>')
    for _, row in df_matrix.iterrows():
        linha = row["Linha"]
        linha_norm = normalizar_texto(linha)
        classe = "result-line" if "resultado contabil" in linha_norm else ("main-line" if linha_norm in [normalizar_texto(x) for x in ["RECEITAS","MARGEM INTERMEDIAÇÃO","RESULTADO ANTES IMPOSTO"]] else "")
        html.append(f'<tr class="{classe}"><td>{linha.replace("*","")}</td>')
        is_pct = linha_pnl_percentual(linha)
        for p in produtos:
            for m in metricas_por_produto[p]:
                val = row[(p, m)]
                cls = []
                if m == "Δ %":
                    if pd.notna(val):
                        val_ex = -val if row[(p, "_ambos_neg")] else val
                        txt = formatar_pontos_percentuais(val_ex) if is_pct else formatar_percentual(val_ex)
                        cls.append("delta-negative" if row[(p, "_delta_bad")] else "delta-positive")
                    else: txt = ""
                elif m == "Δ R$": txt = formatar_numero(val) if pd.notna(val) and not is_pct else ""
                else:
                    txt = formatar_percentual_valor(val) if is_pct else formatar_numero(val)
                    if pd.notna(val) and val < 0: cls.append("neg-value")
                html.append(f'<td class="{" ".join(cls)}">{txt}</td>')
        html.append('</tr>')
    html.append('</tbody></table></div>')
    return "".join(html)

# Lógica de Inicialização e Tabs
st.sidebar.title("Configurações")
upload = st.sidebar.file_uploader("Atualizar base", type=["xlsx"])
arquivo = upload if upload else (Path(ARQUIVO_PADRAO) if Path(ARQUIVO_PADRAO).exists() else None)

if arquivo:
    df_resultado = carregar_resultado(arquivo)
    df_pnl_completo_global = carregar_pnl_mensal(arquivo)
    df_pnl_completo_global = garantir_linha_despesas_administrativas(df_pnl_completo_global)
    
    st.markdown('<div class="dash-title">Dashboard P&L 2026</div>', unsafe_allow_html=True)
    t1, t2, t3, t4 = st.tabs(["Resultados", "P&L Mensal", "P&L Acumulado", "Comparativo 2025"])
    
    with t2: render_pnl_page(df_pnl_completo_global, arquivo, "Mensal")
    with t3: 
        df_c25 = carregar_comparativo_2025(arquivo)
        render_pnl_page(df_pnl_completo_global, arquivo, "Acumulado", df_c25)
    # As outras abas (t1 e t4) seguem a lógica consolidada anteriormente
else:
    st.info("Aguardando arquivo de base...")
