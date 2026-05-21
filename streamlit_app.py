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
        top: 28px;
        left: 22px;
        right: 22px;
    }
    .composition-row {
        margin-bottom: 22px;
    }
    .composition-head {
        display: flex;
        justify-content: space-between;
        gap: 10px;
        margin-bottom: 8px;
        align-items: baseline;
    }
    .composition-name {
        color: #ffffff;
        font-size: 1.05rem;
        font-weight: 700;
    }
    .composition-value {
        color: #ffffff;
        font-size: 1.05rem;
        font-weight: 800;
        text-align: right;
    }
    .composition-pct {
        color: #ffffff;
        font-size: .95rem;
        font-weight: 700;
        min-width: 48px;
        text-align: right;
    }
    .composition-bar-wrap {
        width: 100%;
        height: 10px;
        border-radius: 999px;
        overflow: hidden;
        background: #0b1224;
        border: 1px solid #243150;
    }
    .composition-bar-fill {
        height: 100%;
        border-radius: 999px;
        background: linear-gradient(90deg, #24a8ff 0%, #7cc4ff 100%);
    }
    .composition-help {
        color: #ffffff;
        font-size: .88rem;
        margin-top: 10px;
        line-height: 1.3;
        text-align: center;
    }
    table.pnl-matrix tbody tr.main-line td {
        font-weight: 850 !important;
    }
    table.pnl-matrix tbody tr.main-line td:first-child {
        font-weight: 900 !important;
    }
    table.pnl-matrix tbody tr.main-line td.delta-positive,
    table.pnl-matrix tbody tr.main-line td.delta-negative {
        font-weight: 850 !important;
    }

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

    /* Botão de toggle da sidebar */
    [data-testid="collapsedControl"] {
        color: #ffffff !important;
        background: #0b1224 !important;
        border: 1px solid #1e2a44 !important;
        border-radius: 8px !important;
    }

    .table-wrap {
        width: 100%;
        overflow-x: auto;
        border: 1px solid rgba(255,255,255,.58);
        border-radius: 14px;
        background: #080f1f;
        box-shadow: 0 10px 26px rgba(0,0,0,.20);
    }
    table.dash-table {
        width: 100%;
        border-collapse: collapse;
        background: #080f1f;
        color: #e5ecff;
        font-size: .92rem;
    }
    table.dash-table thead th {
        background: #111a2e;
        color: #ffffff;
        font-weight: 850;
        font-size: 1.05rem;
        text-align: center;
        padding: 15px 14px;
        border-right: 1px solid rgba(255,255,255,.52);
        border-bottom: 1px solid rgba(255,255,255,.70);
        white-space: nowrap;
    }
    table.dash-table thead th:first-child {
        text-align: center;
        min-width: 260px;
    }
    table.dash-table tbody td {
        background: #080f1f;
        color: #e5ecff;
        text-align: center;
        vertical-align: middle;
        padding: 12px 14px;
        border-right: 1px solid rgba(255,255,255,.42);
        border-bottom: 1px solid rgba(255,255,255,.32);
        white-space: nowrap;
    }
    table.dash-table tbody td:first-child {
        color: #ffffff;
        font-weight: 800;
        text-align: left;
    }
    table.dash-table tbody tr:last-child td {
        border-bottom: none;
    }
    table.dash-table th:last-child,
    table.dash-table td:last-child {
        border-right: none;
    }
    table.dash-table tbody tr:hover td {
        background: #111a2e;
    }
    table.dash-table tbody tr.total-row td {
        font-size: 1.04rem;
        font-weight: 900;
        background: #111a2e;
        color: #ffffff;
    }
    table.dash-table tbody tr.total-row td:first-child {
        text-align: center;
    }
    table.dash-table td.neg-value {
        color: #ef4444;
        font-weight: 900;
    }
    table.dash-table td.delta-positive {
        color: #22c55e;
        font-weight: 900;
    }
    table.dash-table td.delta-negative {
        color: #ef4444;
        font-weight: 900;
    }
    table.dash-table td.delta-neutral {
        color: #ffffff;
        font-weight: 850;
    }
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
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
    except Exception:
        return str(valor)
    return f"{valor:,.0f}".replace(",", ".")


def formatar_percentual(valor):
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
    except Exception:
        return str(valor)

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
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
    except Exception:
        return str(valor)

    texto = f"{valor * 100:,.1f}%"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_pontos_percentuais(valor):
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
    except Exception:
        return str(valor)

    sinal = "+" if valor > 0 else ""
    texto = f"{sinal}{valor * 100:,.1f} pp"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_pontos_percentuais_sem_sinal(valor):
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
    except Exception:
        return str(valor)

    texto = f"{abs(valor) * 100:,.1f} pp"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def tabela_html(df, df_valores=None, coluna_delta="Δ mês anterior"):
    html = ['<div class="table-wrap"><table class="dash-table">']
    html.append("<thead><tr>")
    for col in df.columns:
        html.append(f"<th>{col}</th>")
    html.append("</tr></thead><tbody>")

    for idx, row in df.iterrows():
        classe_linha = ' class="total-row"' if str(row.iloc[0]).strip().lower() == "resultado total" else ""
        html.append(f"<tr{classe_linha}>")

        for col in df.columns:
            classes = []

            if col == coluna_delta and df_valores is not None:
                valor_delta = df_valores.loc[idx, col]
                if pd.notna(valor_delta):
                    if valor_delta > 0:
                        classes.append("delta-positive")
                    elif valor_delta < 0:
                        classes.append("delta-negative")
                    else:
                        classes.append("delta-neutral")
            elif col != "Empresa" and df_valores is not None and col in df_valores.columns:
                valor = df_valores.loc[idx, col]
                if pd.notna(valor) and valor < 0:
                    classes.append("neg-value")

            classe_td = f' class="{" ".join(classes)}"' if classes else ""
            html.append(f"<td{classe_td}>{row[col]}</td>")

        html.append("</tr>")

    html.append("</tbody></table></div>")
    return "".join(html)


def converter_periodo(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, pd.Timestamp):
        return valor.to_period("M").to_timestamp()

    if hasattr(valor, "year") and hasattr(valor, "month"):
        try:
            return pd.Timestamp(valor.year, valor.month, 1)
        except Exception:
            pass

    texto = str(valor).strip().lower()
    if not texto or texto == "nan":
        return None

    meses = {
        "jan": 1, "janeiro": 1,
        "fev": 2, "fevereiro": 2,
        "mar": 3, "marco": 3, "março": 3,
        "abr": 4, "abril": 4,
        "mai": 5, "maio": 5,
        "jun": 6, "junho": 6,
        "jul": 7, "julho": 7,
        "ago": 8, "agosto": 8,
        "set": 9, "sep": 9, "setembro": 9,
        "out": 10, "oct": 10, "outubro": 10,
        "nov": 11, "novembro": 11,
        "dez": 12, "dec": 12, "dezembro": 12,
    }

    texto_sem_acento = normalizar_texto(texto)
    partes = texto_sem_acento.split()
    mes = None
    ano = None

    for parte in partes:
        if parte in meses:
            mes = meses[parte]
        elif re.fullmatch(r"\d{4}", parte):
            ano = int(parte)
        elif re.fullmatch(r"\d{2}", parte):
            ano = 2000 + int(parte)

    if mes and ano:
        return pd.Timestamp(ano, mes, 1)

    tentativa = pd.to_datetime(texto, errors="coerce", dayfirst=True)
    if pd.notna(tentativa):
        return tentativa.to_period("M").to_timestamp()

    return None


def nome_periodo(data):
    if pd.isna(data):
        return ""
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

    if valor > 0:
        return "delta-positive"
    if valor < 0:
        return "delta-negative"
    return "delta-neutral"


def card(titulo, valor, ajuda="", variacao=None, variacao_label="Δ mês anterior", cor_classe=None, variacao_exibir=None):
    delta_html = ""
    if variacao is not None:
        cls = cor_classe if cor_classe is not None else classe_variacao(variacao)
        val_txt = variacao_exibir if variacao_exibir is not None else variacao
        delta_html = f'<div class="kpi-delta {cls}">{formatar_variacao(val_txt, variacao_label)}</div>'

    ajuda_html = f'<div class="kpi-help">{ajuda}</div>' if ajuda else ""

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{titulo}</div>
            <div class="kpi-value">{formatar_moeda(valor)}</div>
            {delta_html}
            {ajuda_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def carregar_resultado(arquivo):
    bruto = pd.read_excel(arquivo, sheet_name=ABA_RESULTADO, header=None, engine="openpyxl")
    bruto = bruto.dropna(how="all")

    linha_mes = None
    col_rotulo = None

    for idx in bruto.index:
        for col in bruto.columns:
            if normalizar_texto(bruto.loc[idx, col]) == "mes":
                linha_mes = idx
                col_rotulo = col
                break
        if linha_mes is not None:
            break

    if linha_mes is None or col_rotulo is None:
        raise ValueError("Não encontrei a célula com 'Mês' na aba RESULTADO.")

    colunas_periodo = []
    for col in bruto.columns:
        if col <= col_rotulo:
            continue
        periodo = converter_periodo(bruto.loc[linha_mes, col])
        if periodo is not None:
            colunas_periodo.append((col, periodo))

    if not colunas_periodo:
        raise ValueError("Encontrei a célula 'Mês', mas não encontrei meses válidos na mesma linha.")

    registros = []
    ordem_linha = 0

    for idx in bruto.index:
        if idx <= linha_mes:
            continue

        linha_nome = bruto.loc[idx, col_rotulo]
        if pd.isna(linha_nome) or str(linha_nome).strip() == "":
            continue

        linha_tem_valor = False
        for col, periodo in colunas_periodo:
            valor = pd.to_numeric(bruto.loc[idx, col], errors="coerce")
            if pd.notna(valor):
                linha_tem_valor = True
                registros.append(
                    {
                        "Linha": str(linha_nome).strip(),
                        "Linha_Normalizada": normalizar_texto(linha_nome),
                        "Data": periodo,
                        "Período": nome_periodo(periodo),
                        "Valor": float(valor),
                        "Ordem_Linha": ordem_linha,
                    }
                )

        if linha_tem_valor:
            ordem_linha += 1

    df = pd.DataFrame(registros)

    if df.empty:
        raise ValueError("A aba RESULTADO foi encontrada, mas nenhum valor numérico foi lido.")

    df = df[df["Data"] >= DATA_MINIMA_DASH].copy()

    if df.empty:
        raise ValueError("A aba RESULTADO não possui dados a partir de janeiro/2026.")

    return df


@st.cache_data(show_spinner=False)
def carregar_base_dash(arquivo):
    df = pd.read_excel(arquivo, sheet_name=ABA_BASE, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in ["Visao", "Linha_PnL", "Produto", "Metrica", "Periodo"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    if "Valor" in df.columns:
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0)
    return df


def obter_periodos_pnl_mensal_anualizado(arquivo):
    try:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal - Anualizado", header=None, engine="openpyxl")
    except Exception:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal", header=None, engine="openpyxl")

    periodos = []
    chaves_vistas = set()

    for idx in bruto.index:
        for col in bruto.columns:
            if normalizar_texto(bruto.loc[idx, col]) != "data base":
                continue

            for c_data in range(col + 1, min(col + 12, max(bruto.columns) + 1)):
                if c_data not in bruto.columns:
                    continue

                valor = bruto.loc[idx, c_data]
                data = converter_periodo(valor)

                if data is None:
                    continue

                data_ts = pd.Timestamp(data)

                if data_ts.year < 2020 or data_ts.year > 2035:
                    continue

                chave = data_ts.strftime("%Y-%m")
                if chave not in chaves_vistas:
                    periodos.append({"Período": nome_periodo(data_ts), "Data": data_ts})
                    chaves_vistas.add(chave)

                break

    periodos = sorted(periodos, key=lambda x: x["Data"])

    if not periodos:
        return [{"Período": "Período atual", "Data": None}]

    return periodos


@st.cache_data(show_spinner=False)
def carregar_pnl_mensal(arquivo):
    try:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal - Anualizado", header=None, engine="openpyxl")
    except Exception:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal", header=None, engine="openpyxl")

    bruto = bruto.dropna(how="all")

    registros = []

    for idx in bruto.index:
        for col in bruto.columns:
            if normalizar_texto(bruto.loc[idx, col]) != "data base":
                continue

            linha_data = idx
            linha_produto = idx + 2
            linha_metrica = idx + 3
            col_rotulo = col

            data_base = None
            for c_data in range(col + 1, min(col + 12, max(bruto.columns) + 1)):
                if c_data in bruto.columns:
                    data_base = converter_periodo(bruto.loc[linha_data, c_data])
                    if data_base is not None:
                        break

            if data_base is None:
                continue

            produtos_encontrados = {}
            for c_prod in range(col + 1, min(col + 12, max(bruto.columns) + 1)):
                if c_prod not in bruto.columns:
                    continue

                produto_norm = normalizar_texto(bruto.loc[linha_produto, c_prod])
                if produto_norm in ["consignado", "imobiliario", "total"]:
                    produtos_encontrados[
                        {
                            "consignado": "Consignado",
                            "imobiliario": "Imobiliário",
                            "total": "Total",
                        }[produto_norm]
                    ] = c_prod

            if not {"Consignado", "Imobiliário", "Total"}.issubset(set(produtos_encontrados.keys())):
                continue

            produtos_ordenados = sorted(produtos_encontrados.items(), key=lambda x: x[1])
            blocos = []

            for i, (produto, col_inicio) in enumerate(produtos_ordenados):
                col_fim = produtos_ordenados[i + 1][1] if i + 1 < len(produtos_ordenados) else min(col + 12, max(bruto.columns) + 1)

                for c_met in range(col_inicio, col_fim):
                    if c_met not in bruto.columns:
                        continue

                    metrica_original = bruto.loc[linha_metrica, c_met]
                    metrica_norm = normalizar_texto(metrica_original)

                    if metrica_norm == "realizado":
                        metrica = "Realizado"
                    elif metrica_norm == "orcado":
                        metrica = "Orçado"
                    elif "r" in metrica_norm and ("r" == metrica_norm or "rs" in metrica_norm):
                        metrica = "Δ R$"
                    elif "%" in str(metrica_original) or "delta" in metrica_norm or metrica_norm in [""]:
                        metrica = "Δ %"
                    else:
                        continue

                    blocos.append({"Produto": "Total" if produto == "Total" else produto, "Coluna": c_met, "Métrica": metrica})

            ordem = 0

            for r in bruto.index:
                if r <= linha_metrica:
                    continue

                linha_nome = bruto.loc[r, col_rotulo] if col_rotulo in bruto.columns else None
                if pd.isna(linha_nome) or str(linha_nome).strip() == "":
                    continue

                linha_tem_valor = False

                for bloco in blocos:
                    c_val = bloco["Coluna"]
                    if c_val not in bruto.columns:
                        continue

                    valor = pd.to_numeric(bruto.loc[r, c_val], errors="coerce")
                    if pd.notna(valor):
                        linha_tem_valor = True
                        registros.append(
                            {
                                "Periodo": nome_periodo(data_base),
                                "Data": pd.Timestamp(data_base),
                                "Produto": bloco["Produto"],
                                "Linha": str(linha_nome).strip(),
                                "Linha_Normalizada": normalizar_texto(linha_nome),
                                "Métrica": bloco["Métrica"],
                                "Valor": float(valor),
                                "Ordem_Linha": ordem,
                            }
                        )

                if linha_tem_valor:
                    ordem += 1

    df = pd.DataFrame(registros)

    if df.empty:
        raise ValueError("A aba P&L Mensal - Anualizado foi encontrada, mas nenhum valor numérico foi lido.")

    df = df.drop_duplicates(
        subset=["Periodo", "Produto", "Linha_Normalizada", "Métrica"],
        keep="first",
    ).reset_index(drop=True)

    linhas_em_modulo = [
        normalizar_texto("Alíquota de IR/CSLL"),
        normalizar_texto("Rácio de Eficiência"),
    ]
    mask_modulo = (
        df["Linha_Normalizada"].isin(linhas_em_modulo)
        & df["Métrica"].isin(["Realizado", "Orçado"])
    )
    df.loc[mask_modulo, "Valor"] = df.loc[mask_modulo, "Valor"].abs()

    return df


@st.cache_data(show_spinner=False)
def carregar_pnl_acumulado_oficial_completo(arquivo):
    """Carrega integralmente e de forma direta os valores oficiais da aba 'P&L Acumulado'."""
    try:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Acumulado", header=None, engine="openpyxl")
    except Exception:
        return pd.DataFrame()

    if bruto.empty:
        return pd.DataFrame()

    try:
        data_ate = pd.Timestamp(bruto.iloc[0, 8])
        periodo_label = nome_periodo(data_ate)
    except Exception:
        data_ate = None
        periodo_label = "Acumulado"

    cols_produto = {
        "Consignado": {"Realizado": 5, "Orçado": 6},
        "Imobiliário": {"Realizado": 8, "Orçado": 9},
        "Total": {"Realizado": 11, "Orçado": 12},
    }

    registros = []
    ordem = 0

    for r in bruto.index:
        if r <= 3:  
            continue

        linha_nome = bruto.iloc[r, 1] if 1 < len(bruto.columns) else None
        if pd.isna(linha_nome) or str(linha_nome).strip() == "":
            continue

        linha_tem_valor = False
        linha_norm = normalizar_texto(linha_nome)

        for produto, cols in cols_produto.items():
            col_real = cols["Realizado"]
            col_orc = cols["Orçado"]

            val_real = pd.to_numeric(bruto.iloc[r, col_real], errors="coerce") if col_real < len(bruto.columns) else pd.NA
            val_orc = pd.to_numeric(bruto.iloc[r, col_orc], errors="coerce") if col_orc < len(bruto.columns) else pd.NA

            if pd.notna(val_real):
                linha_tem_valor = True
                registros.append({
                    "Periodo": periodo_label,
                    "Data": data_ate,
                    "Produto": produto,
                    "Linha": str(linha_nome).strip(),
                    "Linha_Normalizada": linha_norm,
                    "Métrica": "Realizado",
                    "Valor": float(val_real),
                    "Ordem_Linha": ordem,
                })
            if pd.notna(val_orc):
                linha_tem_valor = True
                registros.append({
                    "Periodo": periodo_label,
                    "Data": data_ate,
                    "Produto": produto,
                    "Linha": str(linha_nome).strip(),
                    "Linha_Normalizada": linha_norm,
                    "Métrica": "Orçado",
                    "Valor": float(val_orc),
                    "Ordem_Linha": ordem,
                })

        if linha_tem_valor:
            ordem += 1

    df = pd.DataFrame(registros)
    if df.empty:
        return df

    df = df.drop_duplicates(
        subset=["Periodo", "Produto", "Linha_Normalizada", "Métrica"],
        keep="first",
    ).reset_index(drop=True)

    linhas_em_modulo = [
        normalizar_texto("Alíquota de IR/CSLL"),
        normalizar_texto("Rácio de Eficiência"),
    ]
    mask_modulo = (
        df["Linha_Normalizada"].isin(linhas_em_modulo)
        & df["Métrica"].isin(["Realizado", "Orçado"])
    )
    df.loc[mask_modulo, "Valor"] = df.loc[mask_modulo, "Valor"].abs()

    return df


def linhas_ocultas_pnl():
    return {
        normalizar_texto("Componente Juros"),
        normalizar_texto("Componente Inflação"),
        normalizar_texto("RESULTADO ANTES IMPOSTO RECORRENTE"),
        normalizar_texto("Rácio de Eficiência Recorrente"),
        normalizar_texto("Alocação de Capital"),
        normalizar_texto("PL Médio (Banco + Hipo)"),
        normalizar_texto("PL Médio (Prudencial + BRCards)"),
        normalizar_texto("Carteira de Crédito Bruta Média"),
        normalizar_texto("Carteira de Crédito SD Cliente Média"),
        normalizar_texto("Taxa Média Carteira Bruta Média"),
        normalizar_texto("Taxa Média Carteira SD Cliente Média"),
        normalizar_texto("Rateio Carteira"),
        normalizar_texto("Despesas Administrativas"),
    }


def obter_linhas_tabela_pnl(df_pnl):
    if df_pnl.empty:
        return []

    linhas = (
        df_pnl[["Linha", "Linha_Normalizada", "Ordem_Linha"]]
        .drop_duplicates()
        .sort_values("Ordem_Linha")
        .drop_duplicates(subset=["Linha_Normalizada"], keep="first")
    )

    ocultas = linhas_ocultas_pnl()
    linhas = linhas[~linhas["Linha_Normalizada"].isin(ocultas)]

    return linhas["Linha"].tolist()


def obter_linhas_principais_pnl(df_pnl):
    linhas_desejadas = [
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

        if match.empty:
            match = disponiveis[disponiveis["Linha_Normalizada"].str.contains(alvo, regex=False, na=False)]

        if not match.empty:
            selecionadas.append(match.iloc[0]["Linha"])

    return selecionadas


def garantir_linha_despesas_administrativas(df_pnl):
    if df_pnl.empty:
        return df_pnl

    componentes = [
        normalizar_texto("Despesas Administrativas Diretas"),
        normalizar_texto("Desp. Administrativas Indiretas"),
    ]

    base_componentes = df_pnl[df_pnl["Linha_Normalizada"].isin(componentes)].copy()

    if base_componentes.empty:
        mask = pd.Series(False, index=df_pnl.index)
        for comp in componentes:
            mask = mask | df_pnl["Linha_Normalizada"].str.contains(comp, regex=False, na=False)
        base_componentes = df_pnl[mask].copy()

    if base_componentes.empty:
        return df_pnl

    ordem_alvo = df_pnl[df_pnl["Linha_Normalizada"].isin([normalizar_texto("RESULTADO ANTES IMPOSTO")])]["Ordem_Linha"]
    nova_ordem = (ordem_alvo.min() - 0.5) if not ordem_alvo.empty else 99

    cols_grupo = [c for c in ["Produto", "Métrica", "Periodo", "Data"] if c in df_pnl.columns]

    if not cols_grupo:
        return df_pnl

    agregado_base = (
        base_componentes
        .groupby(cols_grupo, as_index=False)["Valor"]
        .sum()
    )

    agregado1 = agregado_base.copy()
    agregado1["Linha"] = "Despesas Administrativas"
    agregado1["Linha_Normalizada"] = normalizar_texto("Despesas Administrativas")
    agregado1["Ordem_Linha"] = nova_ordem

    agregado2 = agregado_base.copy()
    agregado2["Linha"] = "Despesas Adm. Diretas + Indiretas"
    agregado2["Linha_Normalizada"] = normalizar_texto("Despesas Adm. Diretas + Indiretas")
    
    ordem_diretas_indiretas = df_pnl[df_pnl["Linha_Normalizada"] == normalizar_texto("Despesas Adm. Diretas + Indiretas")]["Ordem_Linha"]
    agregado2["Ordem_Linha"] = ordem_diretas_indiretas.min() if not ordem_diretas_indiretas.empty else nova_ordem

    agregado = pd.concat([agregado1, agregado2], ignore_index=True)

    for col in df_pnl.columns:
        if col not in agregado.columns:
            agregado[col] = pd.NA

    agregado = agregado[df_pnl.columns]

    nomes_substituir = {
        normalizar_texto("Despesas Administrativas"),
        normalizar_texto("Despesas Adm. Diretas + Indiretas"),
    }
    df_sem = df_pnl[~df_pnl["Linha_Normalizada"].isin(nomes_substituir)].copy()

    return pd.concat([df_sem, agregado], ignore_index=True)


def valor_pnl(df, produto, linha, metrica):
    base = df[
        (df["Produto"] == produto)
        & (df["Linha"] == linha)
        & (df["Métrica"] == metrica)
    ]

    if base.empty:
        return 0

    if "Ordem_Linha" in base.columns:
        base = base.sort_values("Ordem_Linha")

    return base["Valor"].iloc[0]


def variacao_pnl_mes_anterior(df_pnl_completo, produto, linha, periodo_atual):
    linha_atual = df_pnl_completo[
        (df_pnl_completo["Produto"] == produto)
        & (df_pnl_completo["Linha"] == linha)
        & (df_pnl_completo["Métrica"] == "Realizado")
        & (df_pnl_completo["Periodo"] == periodo_atual)
    ]

    if linha_atual.empty:
        return None

    data_atual = linha_atual["Data"].iloc[0]
    anteriores = (
        df_pnl_completo[
            (df_pnl_completo["Produto"] == produto)
            & (df_pnl_completo["Linha"] == linha)
            & (df_pnl_completo["Métrica"] == "Realizado")
            & (df_pnl_completo["Data"] < data_atual)
        ]
        .sort_values("Data")
    )

    if anteriores.empty:
        return None

    periodo_anterior = anteriores["Periodo"].iloc[-1]

    valor_atual = valor_pnl(df_pnl_completo[df_pnl_completo["Periodo"] == periodo_atual], produto, linha, "Realizado")
    valor_anterior = valor_pnl(
        df_pnl_completo[df_pnl_completo["Periodo"] == periodo_anterior],
        produto,
        linha,
        "Realizado",
    )

    if valor_anterior == 0:
        return None

    return (valor_atual - valor_anterior) / abs(valor_anterior)


def filtrar_pnl_acumulado(df_pnl_completo, periodo_atual):
    linha_periodo = df_pnl_completo[df_pnl_completo["Periodo"] == periodo_atual]

    if linha_periodo.empty:
        return df_pnl_completo.iloc[0:0].copy()

    data_atual = linha_periodo["Data"].iloc[0]
    ano_atual = pd.Timestamp(data_atual).year
    data_inicio = pd.Timestamp(ano_atual, 1, 1)

    base = df_pnl_completo[
        (df_pnl_completo["Data"] >= data_inicio)
        & (df_pnl_completo["Data"] <= data_atual)
    ].copy()

    return base


def linhas_media_pnl():
    return {
        normalizar_texto("Carteira de Crédito Bruta Média"),
        normalizar_texto("Carteira de Crédito SD Cliente Média"),
        normalizar_texto("Carteira de Crédito SD Cliente (média)"),
        normalizar_texto("Carteira de Crédito Média"),
        normalizar_texto("PL Médio"),
        normalizar_texto("PL Médio (Banco + Hipo)"),
        normalizar_texto("PL Médio (Prudencial + BRCards)"),
        normalizar_texto("Alocação de Capital"),
    }


def buscar_linha_acumulada(df_acumulado, produto, inline_norm_alvo, metrica="Realizado"):
    base = df_acumulado[
        (df_acumulado["Produto"] == produto)
        & (df_acumulado["Linha_Normalizada"] == inline_norm_alvo)
        & (df_acumulado["Métrica"] == metrica)
    ]
    if base.empty:
        return None
    return float(base["Valor"].iloc[0])


def recalcular_indicadores_percentuais(df_acumulado, n_meses):
    if df_acumulado.empty or n_meses <= 0:
        return df_acumulado

    fator_anual = 12.0 / n_meses
    produtos = df_acumulado["Produto"].unique()

    novos_registros = []

    for produto in produtos:
        mg_int = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("MARGEM INTERMEDIAÇÃO"))
        mg_int_liq = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("MG INTERMEDIAÇÃO LIQ"))
        res_contabil = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("RESULTADO CONTÁBIL"))
        rai = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("RESULTADO ANTES IMPOSTO"))
        impostos = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Impostos (IR/CSLL)"))
        desp_adm_dir = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Despesas Administrativas Diretas"))
        desp_adm_ind = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Desp. Administrativas Indiretas"))

        carteira_bruta = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Carteira de Crédito Bruta Média"))
        alocacao_capital = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Alocação de Capital"))

        mg_int_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("MARGEM INTERMEDIAÇÃO"), "Orçado")
        mg_int_liq_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("MG INTERMEDIAÇÃO LIQ"), "Orçado")
        res_contabil_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("RESULTADO CONTÁBIL"), "Orçado")
        rai_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("RESULTADO ANTES IMPOSTO"), "Orçado")
        impostos_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Impostos (IR/CSLL)"), "Orçado")
        desp_adm_dir_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Despesas Administrativas Diretas"), "Orçado")
        desp_adm_ind_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Desp. Administrativas Indiretas"), "Orçado")
        carteira_bruta_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Carteira de Crédito Bruta Média"), "Orçado")
        alocacao_capital_orc = buscar_linha_acumulada(df_acumulado, produto, normalizar_texto("Alocação de Capital"), "Orçado")

        def calc_margem_bruta(mg, cart):
            if mg is None or cart is None or cart == 0:
                return None
            return mg * fator_anual / cart

        def calc_margem_liquida(mg_liq, cart):
            if mg_liq is None or cart is None or cart == 0:
                return None
            return mg_liq * fator_anual / cart

        def calc_racio_eficiencia(d_dir, d_ind, mg):
            if d_dir is None or d_ind is None or mg is None or mg == 0:
                return None
            return abs((d_dir + d_ind) / mg)

        def calc_rpl(res, aloc):
            if res is None or aloc is None or aloc == 0:
                return None
            return res * fator_anual / aloc

        def calc_aliquota(imp, rai_v):
            if imp is None or rai_v is None or rai_v == 0:
                return None
            return abs(imp / rai_v)

        indicadores = [
            ("Margem Bruta", normalizar_texto("Margem Bruta"),
             calc_margem_bruta(mg_int, carteira_bruta),
             calc_margem_bruta(mg_int_orc, carteira_bruta_orc)),
            ("Margem Liquida", normalizar_texto("Margem Liquida"),
             calc_margem_liquida(mg_int_liq, carteira_bruta),
             calc_margem_liquida(mg_int_liq_orc, carteira_bruta_orc)),
            ("Rácio de Eficiência", normalizar_texto("Rácio de Eficiência"),
             calc_racio_eficiencia(desp_adm_dir, desp_adm_ind, mg_int),
             calc_racio_eficiencia(desp_adm_dir_orc, desp_adm_ind_orc, mg_int_orc)),
            ("RPL - RES. CONTÁBIL", normalizar_texto("RPL - RES. CONTÁBIL"),
             calc_rpl(res_contabil, alocacao_capital),
             calc_rpl(res_contabil_orc, alocacao_capital_orc)),
            ("Alíquota de IR/CSLL", normalizar_texto("Alíquota de IR/CSLL"),
             calc_aliquota(impostos, rai),
             calc_aliquota(impostos_orc, rai_orc)),
        ]

        for nome_exibir, linha_norm, val_real, val_orc in indicadores:
            ordem_existente = df_acumulado[df_acumulado["Linha_Normalizada"] == linha_norm]["Ordem_Linha"]
            ordem = ordem_existente.iloc[0] if not ordem_existente.empty else 9999

            if val_real is not None:
                novos_registros.append({
                    "Produto": produto,
                    "Linha": nome_exibir,
                    "Linha_Normalizada": linha_norm,
                    "Métrica": "Realizado",
                    "Ordem_Linha": ordem,
                    "Valor": val_real,
                })
            if val_orc is not None:
                novos_registros.append({
                    "Produto": produto,
                    "Linha": nome_exibir,
                    "Linha_Normalizada": linha_norm,
                    "Métrica": "Orçado",
                    "Ordem_Linha": ordem,
                    "Valor": val_orc,
                })

    if not novos_registros:
        return df_acumulado

    linhas_recalculadas = {r["Linha_Normalizada"] for r in novos_registros}
    df_sem_percentuais = df_acumulado[
        ~(df_acumulado["Linha_Normalizada"].isin(linhas_recalculadas)
          & df_acumulado["Métrica"].isin(["Realizado", "Orçado"]))
    ].copy()

    return pd.concat([df_sem_percentuais, pd.DataFrame(novos_registros)], ignore_index=True)


def agregar_pnl_acumulado(df_pnl_periodo):
    if df_pnl_periodo.empty:
        return df_pnl_periodo.copy()

    linhas_percentuais = linhas_percentuais_pnl()
    linhas_media = linhas_media_pnl()

    base_valores = df_pnl_periodo[df_pnl_periodo["Métrica"].isin(["Realizado", "Orçado"])].copy()

    n_meses = base_valores["Periodo"].nunique() if "Periodo" in base_valores.columns else 1
    if n_meses <= 0:
        n_meses = 1

    mask_pct = base_valores["Linha_Normalizada"].isin(linhas_percentuais)
    mask_media = base_valores["Linha_Normalizada"].isin(linhas_media)
    mask_soma = ~(mask_pct | mask_media)

    base_somavel = base_valores[mask_soma]
    base_media = base_valores[mask_media]

    agrupado_soma = (
        base_somavel
        .groupby(["Produto", "Linha", "Linha_Normalizada", "Métrica"], as_index=False)["Valor"]
        .sum()
    )

    if not base_media.empty:
        agrupado_media = (
            base_media
            .groupby(["Produto", "Linha", "Linha_Normalizada", "Métrica"], as_index=False)["Valor"]
            .mean()
        )
    else:
        agrupado_media = base_media.iloc[0:0][["Produto", "Linha", "Linha_Normalizada", "Métrica", "Valor"]]

    agrupado = pd.concat([agrupado_soma, agrupado_media], ignore_index=True)

    ordem_por_linha = (
        base_valores.groupby("Linha_Normalizada", as_index=False)["Ordem_Linha"].min()
    )
    agrupado = agrupado.merge(ordem_por_linha, on="Linha_Normalizada", how="left")

    agrupado = recalcular_indicadores_percentuais(agrupado, n_meses)

    linhas_delta = []

    base_pivot = agrupado.pivot_table(
        index=["Produto", "Linha", "Linha_Normalizada", "Ordem_Linha"],
        columns="Métrica",
        values="Valor",
        aggfunc="first",
    ).reset_index()

    for _, row in base_pivot.iterrows():
        realizado = row.get("Realizado")
        orcado = row.get("Orçado")
        linha_norm = row["Linha_Normalizada"]
        eh_percentual = inline_norm = linha_norm in linhas_percentuais

        if pd.isna(realizado) or pd.isna(orcado):
            delta_rs = pd.NA
            delta_pct = pd.NA
        else:
            delta_rs = realizado - orcado
            if eh_percentual:
                delta_pct = delta_rs
            else:
                delta_pct = pd.NA if orcado == 0 else delta_rs / abs(orcado)

        for metrica, valor in [("Δ %", delta_pct), ("Δ R$", delta_rs)]:
            linhas_delta.append(
                {
                    "Produto": row["Produto"],
                    "Linha": row["Linha"],
                    "Linha_Normalizada": row["Linha_Normalizada"],
                    "Métrica": metrica,
                    "Ordem_Linha": row["Ordem_Linha"],
                    "Valor": valor,
                }
            )

    df_delta = pd.DataFrame(linhas_delta)

    return pd.concat([agrupado, df_delta], ignore_index=True)


def variacao_pnl_acumulado_mes_anterior(df_pnl_completo, produto, linha, periodo_atual):
    linha_atual = df_pnl_completo[
        (df_pnl_completo["Produto"] == produto)
        & (df_pnl_completo["Linha"] == linha)
        & (df_pnl_completo["Métrica"] == "Realizado")
        & (df_pnl_completo["Periodo"] == periodo_atual)
    ]

    if linha_atual.empty:
        return None

    data_atual = linha_atual["Data"].iloc[0]
    ano_atual = pd.Timestamp(data_atual).year
    data_inicio = pd.Timestamp(ano_atual, 1, 1)

    meses_anteriores = (
        df_pnl_completo[
            (df_pnl_completo["Produto"] == produto)
            & (df_pnl_completo["Linha"] == linha)
            & (df_pnl_completo["Métrica"] == "Realizado")
            & (df_pnl_completo["Data"] < data_atual)
            & (df_pnl_completo["Data"] >= data_inicio)
        ]
        .sort_values("Data")
    )

    if meses_anteriores.empty:
        return None

    valor_acumulado_atual = df_pnl_completo[
        (df_pnl_completo["Produto"] == produto)
        & (df_pnl_completo["Linha"] == linha)
        & (df_pnl_completo["Métrica"] == "Realizado")
        & (df_pnl_completo["Data"] >= data_inicio)
        & (df_pnl_completo["Data"] <= data_atual)
    ]["Valor"].sum()

    data_anterior = meses_anteriores["Data"].max()

    valor_acumulado_anterior = df_pnl_completo[
        (df_pnl_completo["Produto"] == produto)
        & (df_pnl_completo["Linha"] == linha)
        & (df_pnl_completo["Métrica"] == "Realizado")
        & (df_pnl_completo["Data"] >= data_inicio)
        & (df_pnl_completo["Data"] <= data_anterior)
    ]["Valor"].sum()

    if valor_acumulado_anterior == 0:
        return None

    return (valor_acumulado_atual - valor_acumulado_anterior) / abs(valor_acumulado_anterior)


def variacao_pnl_acumulado_vs_2025(df_comp_2025, produto, linha, valor_ytd_atual):
    if df_comp_2025 is None or df_comp_2025.empty or pd.isna(valor_ytd_atual):
        return None

    linha_norm = normalizar_texto(linha)
    NORM_DESP_ADM = normalizar_texto("Despesas Administrativas")
    NORM_DESP_DIR = normalizar_texto("Despesas Administrativas Diretas")
    NORM_DESP_IND = normalizar_texto("Desp. Administrativas Indiretas")

    def buscar_valor_2025(norm_alvo):
        base = df_comp_2025[
            (df_comp_2025["Ano"] == 2025)
            & (df_comp_2025["Produto"] == produto)
            & (df_comp_2025["Linha_Normalizada"] == norm_alvo)
        ]
        if base.empty:
            return None
        v = pd.to_numeric(base["Realizado"].iloc[0], errors="coerce")
        return float(v) if pd.notna(v) else None

    if linha_norm == NORM_DESP_ADM:
        v_dir = buscar_valor_2025(NORM_DESP_DIR)
        v_ind = buscar_valor_2025(NORM_DESP_IND)
        if v_dir is None and v_ind is None:
            return None
        valor_2025 = (v_dir or 0.0) + (v_ind or 0.0)
    else:
        valor_2025 = buscar_valor_2025(linha_norm)

    if valor_2025 is None or valor_2025 == 0:
        return None

    return (float(valor_ytd_atual) - valor_2025) / abs(valor_2025)


def render_pnl_page(df_pnl_completo, arquivo, pagina="Mensal", df_comp_2025=None):
    periodos_pnl = obter_periodos_pnl_mensal_anualizado(arquivo)

    if pagina == "Acumulado":
        meses_fim_trimestre = {3, 6, 9, 12}
        periodos_pnl = [p for p in periodos_pnl if p["Data"] is not None and p["Data"].month in meses_fim_trimestre]
        if not periodos_pnl:
            periodos_pnl = obter_periodos_pnl_mensal_anualizado(arquivo)

        mes_para_trimestre = {3: "1Q", 6: "2Q", 9: "3Q", 12: "4Q"}
        label_para_periodo = {}
        lista_labels_trimestre = []
        for p in periodos_pnl:
            if p["Data"] is not None:
                tri = mes_para_trimestre.get(p["Data"].month, p["Período"])
                ano = p["Data"].year
                label = f"{tri} - {ano}"
            else:
                label = p["Período"]
            label_para_periodo[label] = p["Período"]
            lista_labels_trimestre.append(label)

        lista_periodos_pnl = lista_labels_trimestre
    else:
        lista_periodos_pnl = [item["Período"] for item in periodos_pnl]
        label_para_periodo = None

    st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
    
    if pagina == "Acumulado":
        col_produto, col_espaco = st.columns([1, 3.5])
        
        df_pnl = carregar_pnl_acumulado_oficial_completo(arquivo)
        df_pnl = garantir_linha_despesas_administrativas(df_pnl)
        
        if not df_pnl.empty:
            label_sel = df_pnl["Periodo"].iloc[0]
        else:
            label_sel = "Acumulado"
            
        with col_espaco:
            st.info(f"ℹ️ Exibindo dados extraídos diretamente da aba oficial **P&L Acumulado** ({label_sel}).")
    else:
        col_data, col_produto, col_espaco = st.columns([1, 1, 2.5])
        with col_data:
            label_sel = st.selectbox(
                "Data base",
                lista_periodos_pnl,
                index=len(lista_periodos_pnl) - 1,
                key=f"data_pnl_{pagina.lower()}",
            )
            data_sel_pnl = label_para_periodo[label_sel] if label_para_periodo and label_sel in label_para_periodo else label_sel
        df_pnl = df_pnl_completo[df_pnl_completo["Periodo"] == data_sel_pnl].copy()

    empresa_sel_pnl = "Todos"
    opcoes_produto = ["Consignado", "Imobiliário", "Total"]
    index_produto = 2

    with col_produto:
        produto_sel_pnl = st.selectbox(
            "Produto",
            opcoes_produto,
            index=index_produto,
            key=f"produto_pnl_{pagina.lower()}",
        )

    if pagina == "Acumulado":
        titulo_cards = "Principais linhas do P&L Acumulado"
        titulo_comparativo = "Realizado x Orçado acumulado por linha principal"
        titulo_resultado_produto = "Resultado Contábil acumulado por produto"
        titulo_tabela = "Resumo acumulado das linhas principais por produto"
    else:
        titulo_cards = "Principais linhas do P&L Mensal"
        titulo_comparativo = "Realizado x Orçado por linha principal"
        titulo_resultado_produto = "Resultado Contábil por produto"
        titulo_tabela = "Resumo das linhas principais por produto"

    linhas_principais = obter_linhas_principais_pnl(df_pnl)

    st.markdown(f'<div class="section-title">{titulo_cards}</div>', unsafe_allow_html=True)

    for inicio in range(0, len(linhas_principais), 3):
        if inicio > 0:
            st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)

        cols_cards = st.columns(3)
        for col_card, linha in zip(cols_cards, linhas_principais[inicio:inicio + 3]):
            realizado = valor_pnl(df_pnl, produto_sel_pnl, linha, "Realizado")

            if pagina == "Acumulado":
                variacao = variacao_pnl_acumulado_vs_2025(df_comp_2025, produto_sel_pnl, linha, realizado)
                variacao_label = "Δ mesmo período 2025"
            else:
                variacao = variacao_pnl_mes_anterior(df_pnl_completo, produto_sel_pnl, linha, data_sel_pnl)
                variacao_label = "Δ mês anterior"

            cor_classe = None
            variacao_exibir = None
            if variacao is not None and not pd.isna(variacao) and realizado < 0:
                variacao_exibir = -variacao
                cor_classe = "delta-negative" if float(variacao) < 0 else "delta-positive"

            with col_card:
                card_pnl(linha, realizado, variacao=variacao, variacao_label=variacao_label,
                         cor_classe=cor_classe, variacao_exibir=variacao_exibir)

    st.markdown(f'<div class="section-title">{titulo_comparativo}</div>', unsafe_allow_html=True)

    base_grafico = df_pnl[
        (df_pnl["Produto"] == produto_sel_pnl)
        & (df_pnl["Linha"].isin(linhas_principais))
        & (df_pnl["Métrica"].isin(["Realizado", "Orçado"]))
    ].copy()

    ordem_linhas = {linha: i for i, linha in enumerate(linhas_principais)}
    base_grafico["Ordem"] = base_grafico["Linha"].map(ordem_linhas)
    base_grafico = base_grafico.sort_values("Ordem", ascending=False)

    base_grafico["Rótulo"] = base_grafico["Valor"].map(formatar_moeda_curta)

    fig_comp = px.bar(
        base_grafico,
        x="Valor",
        y="Linha",
        color="Métrica",
        text="Rótulo",
        orientation="h",
        barmode="group",
        labels={"Valor": "Valor", "Linha": "", "Métrica": ""},
    )

    if not base_grafico.empty:
        valor_max_abs = base_grafico["Valor"].abs().max()
    else:
        valor_max_abs = 1

    fig_comp.update_traces(
        texttemplate="<b>%{text}</b>",
        textposition="outside",
        textfont=dict(size=14, family="Arial Black", color="#FFFFFF"),
        cliponaxis=False,
        constraintext="none",
    )
    fig_comp.update_layout(
        template="plotly_dark",
        paper_bgcolor="#080f1f",
        plot_bgcolor="#080f1f",
        height=640,
        margin=dict(l=10, r=140, t=20, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=13, color="#ffffff", family="Arial Black")),
        uniformtext_minsize=11,
        uniformtext_mode="show",
        bargap=0.30,
        bargroupgap=0.18,
    )

    if not base_grafico.empty:
        x_min = base_grafico["Valor"].min()
        x_max = base_grafico["Valor"].max()
        x_pad = max((x_max - x_min) * 0.32, 1)
        fig_comp.update_xaxes(
            showgrid=False,
            zeroline=False,
            tickprefix="R$ ",
            separatethousands=True,
            range=[x_min - x_pad, x_max + x_pad],
        )
    else:
        fig_comp.update_xaxes(showgrid=False, zeroline=False, tickprefix="R$ ", separatethousands=True)

    fig_comp.update_yaxes(showgrid=False, zeroline=False, tickfont=dict(size=12, color="#ffffff"))
    st.plotly_chart(fig_comp, use_container_width=True)

    st.markdown(f'<div class="section-title">{titulo_resultado_produto}</div>', unsafe_allow_html=True)
    linha_resultado_contabil = next(
        (linha for linha in linhas_principais if normalizar_texto(linha) in ["resultado contabil", "resultado contábil"]),
        linhas_principais[-1] if linhas_principais else None,
    )

    base_produtos = df_pnl[
        (df_pnl["Linha"] == linha_resultado_contabil)
        & (df_pnl["Produto"].isin(["Consignado", "Imobiliário", "Total"]))
        & (df_pnl["Métrica"] == "Realizado")
    ].copy()

    def texto_barra(row):
        val_str = formatar_moeda(row["Valor"])

        if pagina == "Acumulado":
            if df_comp_2025 is not None and not df_comp_2025.empty:
                linha_norm = normalizar_texto(linha_resultado_contabil)
                b26 = df_comp_2025[
                    (df_comp_2025["Ano"] == 2026)
                    & (df_comp_2025["Produto"] == row["Produto"])
                    & (df_comp_2025["Linha_Normalizada"] == linha_norm)
                ]
                b25 = df_comp_2025[
                    (df_comp_2025["Ano"] == 2025)
                    & (df_comp_2025["Produto"] == row["Produto"])
                    & (df_comp_2025["Linha_Normalizada"] == linha_norm)
                ]
                v26 = pd.to_numeric(b26["Realizado"].iloc[0], errors="coerce") if not b26.empty else None
                v25 = pd.to_numeric(b25["Realizado"].iloc[0], errors="coerce") if not b25.empty else None
                if v26 is not None and v25 is not None and pd.notna(v26) and pd.notna(v25) and v25 != 0:
                    var = (float(v26) / float(v25)) - 1
                else:
                    var = None
            else:
                var = None
            label_var = "vs 1Q25"
        else:
            var = variacao_pnl_mes_anterior(df_pnl_completo, row["Produto"], linha_resultado_contabil, data_sel_pnl)
            label_var = "vs mês ant."

        if var is None or pd.isna(var):
            return val_str

        sinal = "+" if float(var) >= 0 else "−"
        pct = f"{abs(float(var)) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{val_str}<br><span style='font-size:14px'>{sinal}{pct} {label_var}</span>"

    base_produtos["Rótulo"] = base_produtos.apply(texto_barra, axis=1)

    max_val = base_produtos["Valor"].max() if not base_produtos.empty else 1
    threshold = max_val * 0.20

    text_positions = [
        "outside" if row["Valor"] < threshold else "inside"
        for _, row in base_produtos.iterrows()
    ]

    fig_prod = go.Figure(go.Bar(
        x=base_produtos["Produto"],
        y=base_produtos["Valor"],
        text=base_produtos["Rótulo"],
        textposition=text_positions,
        textfont=dict(size=16, family="Arial Black", color="#FFFFFF"),
        insidetextanchor="middle",
        texttemplate="%{text}",
        marker_color="#1d6ff2",
    ))
    fig_prod.update_layout(
        template="plotly_dark",
        paper_bgcolor="#080f1f",
        plot_bgcolor="#080f1f",
        height=390,
        margin=dict(l=10, r=10, t=10, b=10),
        showlegend=False,
    )
    fig_prod.update_xaxes(showgrid=False, zeroline=False)
    fig_prod.update_yaxes(showgrid=False, zeroline=False, tickprefix="R$ ", separatethousands=True)
    st.plotly_chart(fig_prod, use_container_width=True)

    st.markdown(f'<div class="section-title">{titulo_tabela}</div>', unsafe_allow_html=True)
    linhas_tabela = obter_linhas_tabela_pnl(df_pnl)
    matriz_pnl, produtos_matriz, metricas_matriz = montar_matriz_pnl_excel(df_pnl, linhas_tabela)
    st.markdown(
        tabela_html_pnl_matriz(matriz_pnl, produtos_matriz, metricas_matriz),
        unsafe_allow_html=True,
    )


def card_pnl(titulo, valor, variacao=None, variacao_label="Δ mês anterior", cor_classe=None, variacao_exibir=None):
    if variacao is None or pd.isna(variacao):
        delta_html = '<div class="kpi-delta delta-neutral">N/D</div>'
    else:
        cls = cor_classe if cor_classe is not None else classe_variacao(variacao)
        val_txt = variacao_exibir if variacao_exibir is not None else variacao
        delta_html = f'<div class="kpi-delta {cls}">{formatar_variacao(val_txt, variacao_label)}</div>'

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{titulo}</div>
            <div class="kpi-value">{formatar_moeda(valor)}</div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def montar_tabela_pnl_principal(df_pnl, linhas_principais):
    base = df_pnl[df_pnl["Linha"].isin(linhas_principais)].copy()

    tabela = (
        base.pivot_table(
            index=["Produto", "Linha"],
            columns="Métrica",
            values="Valor",
            aggfunc="sum",
        )
        .reset_index()
    )

    for col in ["Realizado", "Orçado", "Δ %", "Δ R$"]:
        if col not in tabela.columns:
            tabela[col] = pd.NA

    ordem_linhas = {linha: i for i, linha in enumerate(linhas_principais)}
    ordem_produtos = {"Consignado": 1, "Imobiliário": 2, "Total": 3}
    tabela["ordem_linha"] = tabela["Linha"].map(ordem_linhas)
    tabela["ordem_produto"] = tabela["Produto"].map(ordem_produtos)
    tabela = tabela.sort_values(["ordem_produto", "ordem_linha"]).drop(columns=["ordem_produto", "ordem_linha"])

    return tabela[["Produto", "Linha", "Realizado", "Orçado", "Δ %", "Δ R$"]]


def tabela_html_pnl(df, df_valores=None):
    html = ['<div class="table-wrap"><table class="dash-table">']
    html.append("<thead><tr>")
    for col in df.columns:
        html.append(f"<th>{col}</th>")
    html.append("</tr></thead><tbody>")

    for idx, row in df.iterrows():
        is_total = str(row.get("Linha", "")).strip().lower() in ["resultado contábil", "resultado contabil"]
        classe_linha = ' class="total-row"' if is_total else ""
        html.append(f"<tr{classe_linha}>")

        for col in df.columns:
            classes = []
            if df_valores is not None and col in df_valores.columns and col not in ["Produto", "Linha"]:
                valor = df_valores.loc[idx, col]
                if pd.notna(valor) and isinstance(valor, (int, float)) and valor < 0:
                    classes.append("neg-value")

            classe_td = f' class="{" ".join(classes)}"' if classes else ""
            html.append(f"<td{classe_td}>{row[col]}</td>")

        html.append("</tr>")

    html.append("</tbody></table></div>")
    return "".join(html)


def montar_matriz_pnl_excel(df_pnl, linhas_principais):
    produtos = ["Consignado", "Imobiliário", "Total"]
    metricas_por_produto = {
        "Consignado": ["Realizado", "Orçado", "Δ %"],
        "Imobiliário": ["Realizado", "Orçado", "Δ %"],
        "Total": ["Realizado", "Orçado", "Δ %", "Δ R$"],
    }

    linhas = []

    for linha in linhas_principais:
        row = {"Linha": inline_ref_label := linha} # simple tracking anchor trace
        linha_percentual = linha_pnl_percentual(linha)
        linha_norm_card = normalizar_texto(linha)

        racio_eficiencia = linha_norm_card == normalizar_texto("Rácio de Eficiência")

        for produto in produtos:
            realizado = valor_pnl(df_pnl, produto, linha, "Realizado")
            orcado = valor_pnl(df_pnl, produto, linha, "Orçado")

            if racio_eficiencia:
                delta_rs = orcado - realizado
            else:
                delta_rs = realizado - orcado

            if linha_percentual:
                delta_pct = delta_rs if pd.notna(delta_rs) else pd.NA
            else:
                delta_pct = pd.NA if orcado == 0 else delta_rs / abs(orcado)

            row[(produto, "Realizado")] = realizado
            row[(produto, "Orçado")] = orcado
            row[(produto, "Δ %")] = delta_pct

            ambos_neg = (
                not racio_eficiencia
                and not linha_percentual
                and pd.notna(realizado) and pd.notna(orcado)
                and realizado < 0 and orcado < 0
            )
            row[(produto, "_ambos_neg")] = ambos_neg

            if pd.isna(realizado) or pd.isna(orcado) or pd.isna(delta_pct):
                delta_bad = False
            elif racio_eficiencia:
                delta_bad = False
            else:
                valor_exibido = -float(delta_pct) if ambos_neg else float(delta_pct)
                delta_bad = valor_exibido < 0

            row[(produto, "_delta_bad")] = delta_bad

            if produto == "Total":
                row[(produto, "Δ R$")] = pd.NA if linha_percentual else delta_rs

        linhas.append(row)

    return pd.DataFrame(linhas), produtos, metricas_por_produto


def tabela_html_pnl_matriz(df_matrix, produtos, metricas_por_produto):
    linhas_destaque = {
        normalizar_texto("RECEITAS"),
        normalizar_texto("Operações de Crédito"),
        normalizar_texto("DESPESAS DE ORIGINAÇÃO"),
        normalizar_texto("MARGEM INTERMEDIAÇÃO"),
        normalizar_texto("MG INTERMEDIAÇÃO LIQ"),
        normalizar_texto("MG CONTRIBUIÇÃO DIRETA"),
        normalizar_texto("RESULTADO ANTES IMPOSTO"),
        normalizar_texto("RESULTADO CONTÁBIL"),
    }

    html = ['<div class="table-wrap"><table class="pnl-matrix">']

    html.append("<thead>")
    html.append("<tr>")
    html.append('<th rowspan="2">Linha P&L</th>')
    for produto in produtos:
        html.append(f'<th class="product-header" colspan="{len(metricas_por_produto[produto])}">{produto.upper()}</th>')
    html.append("</tr>")

    html.append("<tr>")
    for produto in produtos:
        for metrica in metricas_por_produto[produto]:
            html.append(f'<th class="metric-header">{metrica}</th>')
    html.append("</tr>")
    html.append("</thead><tbody>")

    for _, row in df_matrix.iterrows():
        linha = row["Linha"]
        linha_norm = normalizar_texto(linha)

        if linha_norm in ["resultado contabil", "resultado contábil"]:
            classe = "result-line"
        elif linha_norm in linhas_destaque:
            classe = "main-line"
        else:
            classe = ""

        tr_class = f' class="{classe}"' if classe else ""
        html.append(f"<tr{tr_class}>")
        linha_display = linha.replace("*", "").strip()
        html.append(f"<td>{linha_display}</td>")

        linha_percentual = inline_pct_check = linha_pnl_percentual(linha)
        ocultar_variacao = linha_norm == normalizar_texto("Alíquota de IR/CSLL")
        eh_racio_eficiencia = linha_norm == normalizar_texto("Rácio de Eficiência")

        for produto in produtos:
            for metrica in metricas_por_produto[produto]:
                valor = row[(produto, metrica)]
                classes = []

                if metrica == "Δ %":
                    if ocultar_variacao:
                        texto = ""
                    elif eh_racio_eficiencia:
                        texto = formatar_pontos_percentuais(valor)
                        if pd.notna(valor):
                            classes.append("delta-positive" if valor >= 0 else "delta-negative")
                    else:
                        if pd.notna(valor):
                            ambos_neg = row[(produto, "_ambos_neg")]
                            valor_exibir = -valor if ambos_neg else valor
                            texto = formatar_pontos_percentuais(valor_exibir) if linha_percentual else formatar_percentual(valor_exibir)
                            classes.append("delta-negative" if row[(produto, "_delta_bad")] else "delta-positive")
                        else:
                            texto = formatar_pontos_percentuais(valor) if linha_percentual else formatar_percentual(valor)
                elif metrica == "Δ R$":
                    if ocultar_variacao:
                        texto = ""
                    else:
                        if pd.notna(valor) and not linha_percentual:
                            ambos_neg = row[(produto, "_ambos_neg")]
                            valor_exibir = -valor if ambos_neg else valor
                            texto = formatar_numero(valor_exibir)
                            classes.append("delta-negative" if row[(produto, "_delta_bad")] else "delta-positive")
                        else:
                            texto = "" if linha_percentual else formatar_numero(valor)
                elif linha_percentual and metrica in ["Realizado", "Orçado"]:
                    texto = formatar_percentual_valor(valor)
                    if pd.notna(valor) and valor < 0:
                        classes.append("neg-value")
                else:
                    texto = formatar_numero(valor)
                    if pd.notna(valor) and valor < 0:
                        classes.append("neg-value")

                classe_td = f' class="{" ".join(classes)}"' if classes else ""
                html.append(f"<td{classe_td}>{texto}</td>")

        html.append("</tr>")

    html.append("</tbody></table></div>")
    return "".join(html)


def achar_linha_exata_ou_contendo(df, termos):
    linhas = df[["Linha", "Linha_Normalizada", "Ordem_Linha"]].drop_duplicates().sort_values("Ordem_Linha")
    for termo in termos:
        termo_norm = normalizar_texto(termo)
        exato = linhas[linhas["Linha_Normalizada"].eq(termo_norm)]
        if not exato.empty:
            return exato.iloc[0]["Linha"]
    for termo in termos:
        termo_norm = normalizar_texto(termo)
        encontrado = linhas[linhas["Linha_Normalizada"].str.contains(termo_norm, regex=False, na=False)]
        if not encontrado.empty:
            return encontrado.iloc[0]["Linha"]
    return None


def montar_resultados_principais(df):
    specs = [
        ("Resultado Conglomerado Financeiro", ["resultado congl financeiro", "resultado conglomerado financeiro"]),
        ("Resultado Coligadas", ["resultado coligadas"]),
        ("Resultado Conglomerado + Coligadas", ["resultado congl coligadas", "resultado conglomerado coligadas"]),
        ("Resultado Total", ["res total", "resultado total"]),
    ]
    mapeamento = []
    for titulo, termos in specs:
        linha = achar_linha_exata_ou_contendo(df, termos)
        if linha:
            mapeamento.append({"Indicador": titulo, "Linha": linha})

    if not mapeamento:
        return pd.DataFrame(columns=["Indicador", "Linha", "Data", "Período", "Valor"])

    mapa = pd.DataFrame(mapeamento)
    return df.merge(mapa, on="Linha", how="inner")


def periodo_anterior(periodos_df, periodo_atual):
    linha_atual = periodos_df[periodos_df["Período"] == periodo_atual]
    if linha_atual.empty:
        return None

    data_atual = linha_atual["Data"].iloc[0]
    anteriores = periodos_df[periodos_df["Data"] < data_atual].sort_values("Data")

    if anteriores.empty:
        return None

    return anteriores.iloc[-1]["Período"]


def variacao_mes_anterior(df_principais, indicador, periodo_atual, periodo_ant):
    if periodo_ant is None:
        return None

    valor_atual = df_principais[
        (df_principais["Indicador"] == indicador)
        & (df_principais["Período"] == periodo_atual)
    ]["Valor"].sum()

    valor_ant = df_principais[
        (df_principais["Indicador"] == indicador)
        & (df_principais["Período"] == periodo_ant)
    ]["Valor"].sum()

    if valor_ant == 0:
        return None

    return (valor_atual - valor_ant) / abs(valor_ant)


def resultado_total_acumulado_ano(df_principais, periodo_atual):
    linha_atual = df_principais[
        (df_principais["Indicador"] == "Resultado Total")
        & (df_principais["Período"] == periodo_atual)
    ]

    if linha_atual.empty:
        return None, None, None, None

    data_atual = linha_atual["Data"].iloc[0]
    ano_atual = pd.Timestamp(data_atual).year
    data_inicio = pd.Timestamp(ano_atual, 1, 1)

    base_ano = df_principais[
        (df_principais["Indicador"] == "Resultado Total")
        & (df_principais["Data"] >= data_inicio)
        & (df_principais["Data"] <= data_atual)
    ].copy()

    if base_ano.empty:
        return None, None, None, None

    valor_acumulado = base_ano["Valor"].sum()

    data_mes_anterior = pd.Timestamp(data_atual) - pd.DateOffset(months=1)
    base_ate_mes_anterior = base_ano[base_ano["Data"] <= data_mes_anterior]

    valor_acumulado_anterior = base_ate_mes_anterior["Valor"].sum() if not base_ate_mes_anterior.empty else None

    if valor_acumulado_anterior is None or valor_acumulado_anterior == 0:
        variacao = None
    else:
        variacao = (valor_acumulado - valor_acumulado_anterior) / abs(valor_acumulado_anterior)

    return valor_acumulado, variacao, valor_acumulado_anterior, data_inicio


def card_resultado_total_acumulado(valor_acumulado, variacao, valor_acumulado_anterior, periodo_atual):
    if valor_acumulado is None:
        valor_html = "N/D"
        ajuda = "Resultado Total não encontrado"
    else:
        valor_html = formatar_moeda(valor_acumulado)
        ajuda = f"Acumulado de jan/2026 até {periodo_atual}"

    st.markdown(
        f"""
        <div class="side-card">
            <div class="side-card-label">Resultado Total acumulado em 2026</div>
            <div class="side-card-value">{valor_html}</div>
            <div class="side-card-help">{ajuda}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def composicao_resultado_total_acumulado_produto(df_pnl_completo, periodo_atual, empresa_sel="Todos"):
    if df_pnl_completo is None or df_pnl_completo.empty:
        return None, []

    df_periodo = filtrar_pnl_acumulado(df_pnl_completo, periodo_atual)
    df_acumulado = agregar_pnl_acumulado(df_periodo)
    if df_acumulado.empty:
        return None, []

    linhas_principais = obter_linhas_principais_pnl(df_acumulado)
    
    linha_resultado_contabil = next(
        (linha for linha in linhas_principais if normalizar_texto(linha) in ["resultado contabil", "resultado contábil"]),
        None,
    )

    if linha_resultado_contabil is None:
        candidatos = df_acumulado[df_acumulado["Linha_Normalizada"].str.contains("resultado contabil", na=False, regex=False)]
        if not candidatos.empty:
            linha_resultado_contabil = candidatos.sort_values("Ordem_Linha").iloc[0]["Linha"]

    if linha_resultado_contabil is None:
        return None, []

    valor_total = valor_pnl(df_acumulado, "Total", linha_resultado_contabil, "Realizado")
    valor_consignado = valor_pnl(df_acumulado, "Consignado", linha_resultado_contabil, "Realizado")
    valor_imobiliario = valor_pnl(df_acumulado, "Imobiliário", linha_resultado_contabil, "Realizado")
    valor_ajustes = valor_total - (valor_consignado + valor_imobiliario)

    if empresa_sel == "Banco":
        componentes = [("Consignado", valor_consignado)]
        total_base = valor_consignado
    elif empresa_sel == "Hipotecária":
        componentes = [("Imobiliário", valor_imobiliario)]
        total_base = valor_imobiliario
    else:
        componentes = [("Consignado", valor_consignado), ("Imobiliário", valor_imobiliario)]
        if abs(valor_ajustes) > 0.5:
            componentes.append(("Ajustes / Outros", valor_ajustes))
        total_base = valor_total

    itens = []
    base_pct = abs(total_base) if total_base not in (None, 0) else None
    for nome, valor in componentes:
        pct = (valor / base_pct) if base_pct else None
        itens.append({"nome": nome, "valor": valor, "pct": pct})

    return total_base, itens


def card_composicao_resultado_total_acumulado(df_pnl_completo, periodo_atual, empresa_sel="Todos"):
    total, itens = composicao_resultado_total_acumulado_produto(df_pnl_completo, periodo_atual, empresa_sel)

    if total is None or not itens:
        html = (
            '<div class="side-card composition-card">'
            '<div class="composition-title">Composição do Resultado Total acumulado</div>'
            '<div class="composition-help">Não foi possível calcular a composição para o período selecionado.</div>'
            '</div>'
        )
        st.markdown(html, unsafe_allow_html=True)
        return

    html_rows = []
    for item in itens:
        pct = item["pct"] if item["pct"] is not None else 0.0
        pct_texto = f"{pct * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")
        largura = max(3, abs(pct) * 100)
        row_html = (
            '<div class="composition-row">'
            '<div class="composition-head">'
            f'<div class="composition-name">{item["nome"]}</div>'
            '<div style="display:flex; gap:8px; align-items:baseline;">'
            f'<div class="composition-value">{formatar_moeda(item["valor"])}</div>'
            f'<div class="composition-pct">{pct_texto}</div>'
            '</div>'
            '</div>'
            '<div class="composition-bar-wrap">'
            f'<div class="composition-bar-fill" style="width:{largura:.1f}%"></div>'
            '</div>'
            '</div>'
        )
        html_rows.append(row_html)

    ajuda = f"Composição do acumulado de jan/2026 até {periodo_atual}"
    html = (
        '<div class="side-card composition-card">'
        '<div class="composition-title">Composição do Resultado Total acumulado</div>'
        + ''.join(html_rows)
        + f'<div class="composition-help">{ajuda}</div>'
        + '</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


def filtrar_tabela_resultado_por_empresa(tabela, empresa_sel):
    if empresa_sel == "Todos" or table_check := tabela.empty: # standard variable tracking logic check
        return tabela

    col_nome = tabela.columns[0]
    base = tabela.copy()
    base["_nome_norm"] = base[col_nome].astype(str).map(normalizar_texto)

    banco_set = {"banco", "equiv patr", "jcp dividendos", "resultado banco"}

    if empresa_sel == "Banco":
        filtrada = base[base["_nome_norm"].isin(banco_set) | base["_nome_norm"].str.contains("banco", regex=False, na=False)]
    else:
        filtrada = base[
            ~base["_nome_norm"].isin(banco_set)
            & ~base["_nome_norm"].eq("resultado total")
        ]

    return filtrada.drop(columns=["_nome_norm"])


def adicionar_coluna_variacao_tabela(tabela, periodos_df, periodo_atual):
    coluna_delta = "Δ mês anterior"
    tabela = tabela.copy()

    periodo_ant = periodo_anterior(periodos_df, periodo_atual)

    if periodo_ant is None or periodo_ant not in tabela.columns or periodo_atual not in tabela.columns:
        tabela[coluna_delta] = pd.NA
        return tabela, coluna_delta

    atual = pd.to_numeric(tabela[periodo_atual], errors="coerce")
    anterior = pd.to_numeric(tabela[periodo_ant], errors="coerce")

    tabela[coluna_delta] = (atual - anterior) / anterior.abs()
    tabela.loc[anterior.eq(0) | anterior.isna(), coluna_delta] = pd.NA

    return tabela, coluna_delta


def linha_principal_comparativo(linha):
    linhas = {
        normalizar_texto("RECEITAS"),
        normalizar_texto("Operações de Crédito"),
        normalizar_texto("DESPESAS DE ORIGINAÇÃO"),
        normalizar_texto("DESPESAS TOTAIS"),
        normalizar_texto("Provisões"),
        normalizar_texto("MARGEM INTERMEDIAÇÃO"),
        normalizar_texto("MG INTERMEDIAÇÃO LIQ"),
        normalizar_texto("MG CONTRIBUIÇÃO DIRETA"),
        normalizar_texto("RESULTADO ANTES IMPOSTO"),
        normalizar_texto("RESULTADO CONTÁBIL"),
    }
    return normalizar_texto(linha) in linhas


def carregar_comparativo_2025(arquivo):
    try:
        bruto = pd.read_excel(arquivo, sheet_name="Comparativo 2026 x 2025", header=None, engine="openpyxl")
    except Exception:
        bruto = pd.read_excel(arquivo, sheet_name="Comparativo 2025", header=None, engine="openpyxl")

    label_cols = []
    for col in bruto.columns:
        for r in range(min(50, len(bruto))):
            val = normalizar_texto(bruto.iat[r, col])
            if val in ["receitas", "resultado contabil", "resultado contábil", "despesas totais"]:
                if col not in label_cols:
                    label_cols.append(col)
                break

    if not label_cols:
        label_cols = [1, 15]

    col_26 = label_cols[0]
    col_25 = label_cols[1] if len(label_cols) > 1 else None

    def extrair_bloco(label_col, ano):
        registros_bloco = []
        if label_col is None or label_col >= len(bruto.columns):
            return registros_bloco

        blocos = [
            {"Produto": "Consignado",  "realizado_col": label_col + 4,  "orcado_col": label_col + 5},
            {"Produto": "Imobiliário", "realizado_col": label_col + 7,  "orcado_col": label_col + 8},
            {"Produto": "Total",       "realizado_col": label_col + 10, "orcado_col": label_col + 11},
        ]

        for bloco in blocos:
            for idx_row in bruto.index:
                linha = bruto.iat[idx_row, label_col]
                linha_norm = normalizar_texto(linha)
                if not linha_norm:
                    continue
                
                realizado = pd.NA
                if bloco["realizado_col"] < len(bruto.columns):
                    try:
                        v = bruto.iat[idx_row, bloco["realizado_col"]]
                        if pd.notna(v): realizado = float(v)
                    except: pass
                    
                orcado = pd.NA
                if bloco["orcado_col"] < len(bruto.columns):
                    try:
                        v = bruto.iat[idx_row, bloco["orcado_col"]]
                        if pd.notna(v): orcado = float(v)
                    except: pass
                    
                registros_bloco.append({
                    "Ano": ano,
                    "Produto": bloco["Produto"],
                    "Linha": str(linha).strip(),
                    "Linha_Normalizada": linha_norm,
                    "Realizado": realizado,
                    "Orçado": orcado,
                    "Ordem": int(idx_row),
                })
        return registros_bloco

    registros = []
    registros.extend(extrair_bloco(col_26, 2026))
    if col_25 is not None:
        registros.extend(extrair_bloco(col_25, 2025))

    df = pd.DataFrame(registros)
    if df.empty:
        return df

    return df.sort_values(["Ano", "Produto", "Ordem"]).drop_duplicates(
        ["Ano", "Produto", "Linha_Normalizada"], keep="first"
    )


def carregar_2025_acumulado(arquivo):
    try:
        bruto = pd.read_excel(arquivo, sheet_name="2025 Acumulado", header=None, engine="openpyxl")
    except Exception:
        return pd.DataFrame()

    label_col = 0
    for col in bruto.columns:
        for r in range(min(50, len(bruto))):
            val = normalizar_texto(bruto.iat[r, col])
            if val in ["receitas", "resultado contabil", "resultado contábil"]:
                label_col = col
                break
        if label_col != 0:
            break

    col_real = label_col + 7
    col_orc = label_col + 8
    col_delta = label_col + 9

    def valor_numero(v):
        if pd.isna(v): return pd.NA
        try: return float(v)
        except: return pd.NA

    registros = []
    for idx_row in bruto.index:
        linha = bruto.iat[idx_row, label_col] if label_col < len(bruto.columns) else None
        linha_norm = normalizar_texto(linha)
        if not linha_norm:
            continue

        realizado_total = valor_numero(bruto.iat[idx_row, col_real]) if col_real < len(bruto.columns) else pd.NA
        orcado_total = valor_numero(bruto.iat[idx_row, col_orc]) if col_orc < len(bruto.columns) else pd.NA
        delta_total = valor_numero(bruto.iat[idx_row, col_delta]) if col_delta < len(bruto.columns) else pd.NA

        registros.append({
            "Linha": str(linha).strip(),
            "Linha_Normalizada": linha_norm,
            "Realizado": realizado_total,
            "Orçado": orcado_total,
            "Δ Orçado": delta_total,
            "Ordem": int(idx_row),
        })

    df = pd.DataFrame(registros)
    if df.empty:
        return df

    return df.sort_values("Ordem").drop_duplicates(["Linha_Normalizada"], keep="first")


def montar_comparativo_principais(df_comp, df_2025_acumulado=None):
    # === ADICIONADAS LINHAS DE MÉDIAS AO MAPEAMENTO DE COMPARAÇÃO ===
    linhas_ordem = [
        "RECEITAS",
        "Operações de Crédito",
        "DESPESAS DE ORIGINAÇÃO",
        "DESPESAS TOTAIS",
        "Provisões",
        "MARGEM INTERMEDIAÇÃO",
        "MG INTERMEDIAÇÃO LIQ",
        "MG CONTRIBUIÇÃO DIRETA",
        "RESULTADO ANTES IMPOSTO",
        "RESULTADO CONTÁBIL",
        "Carteira de Crédito Bruta Média",
        "Carteira de Crédito Média",
        "PL Médio",
        "PL Médio (Banco + Hipo)",
    ]

    componentes_desp_totais = [
        "despesas de originacao",
        "provisoes",
        "despesas administrativas",
        "despesas administrativas diretas",
        "desp administrativas indiretas",
        "amortizacao estoque projeto banco",
    ]

    if df_2025_acumulado is None:
        df_2025_acumulado = pd.DataFrame()

    def somar_componentes_desp(df_ano, ano, acumulado_df=None):
        soma_26 = 0.0
        soma_25 = 0.0
        soma_25a = 0.0
        encontrou_26 = False
        encontrou_25 = False
        encontrou_25a = False

        for norm_comp in componentes_desp_totais:
            b26 = df_comp[(df_comp["Ano"] == 2026) & (df_comp["Produto"] == "Total") & (df_comp["Linha_Normalizada"].str.contains(norm_comp, na=False, regex=False))]
            b25 = df_comp[(df_comp["Ano"] == 2025) & (df_comp["Produto"] == "Total") & (df_comp["Linha_Normalizada"].str.contains(norm_comp, na=False, regex=False))]

            if not b26.empty and pd.notna(b26["Realizado"].iloc[0]):
                soma_26 += b26["Realizado"].iloc[0]
                encontrou_26 = True
            if not b25.empty and pd.notna(b25["Realizado"].iloc[0]):
                soma_25 += b25["Realizado"].iloc[0]
                encontrou_25 = True

            if not acumulado_df.empty:
                ba = acumulado_df[acumulado_df["Linha_Normalizada"].str.contains(norm_comp, na=False, regex=False)]
                if not ba.empty and pd.notna(ba["Realizado"].iloc[0]):
                    soma_25a += ba["Realizado"].iloc[0]
                    encontrou_25a = True

        return (
            soma_26 if encontrou_26 else pd.NA,
            soma_25 if encontrou_25 else pd.NA,
            soma_25a if encontrou_25a else pd.NA,
        )

    linhas = []
    for ordem, linha_ref in enumerate(linhas_ordem):
        linha_norm = normalizar_texto(linha_ref)

        if linha_norm == normalizar_texto("DESPESAS TOTAIS"):
            v26, v25, v25_acum = somar_componentes_desp(df_comp, 2026, df_2025_acumulado)
        else:
            b25 = df_comp[(df_comp["Ano"] == 2025) & (df_comp["Produto"] == "Total") & (df_comp["Linha_Normalizada"] == linha_norm)]
            b26 = df_comp[(df_comp["Ano"] == 2026) & (df_comp["Produto"] == "Total") & (df_comp["Linha_Normalizada"] == linha_norm)]
            b25_acum = (
                df_2025_acumulado[df_2025_acumulado["Linha_Normalizada"] == linha_norm]
                if not df_2025_acumulado.empty and "Linha_Normalizada" in df_2025_acumulado.columns
                else pd.DataFrame()
            )
            v25 = b25["Realizado"].iloc[0] if not b25.empty else pd.NA
            v26 = b26["Realizado"].iloc[0] if not b26.empty else pd.NA
            v25_acum = b25_acum["Realizado"].iloc[0] if not b25_acum.empty else pd.NA

        delta_rs = v26 - v25 if pd.notna(v25) and pd.notna(v26) else pd.NA
        delta_pct = delta_rs / abs(v25) if pd.notna(delta_rs) and v25 not in [0, 0.0] and pd.notna(v25) else pd.NA
        alcance = v26 / abs(v25_acum) if pd.notna(v25_acum) and v25_acum not in [0, 0.0] and pd.notna(v26) else pd.NA

        ambos_neg = pd.notna(v25) and pd.notna(v26) and float(v25) < 0 and float(v26) < 0
        if ambos_neg:
            delta_bad = abs(float(v26)) > abs(float(v25)) if pd.notna(delta_rs) else False
        elif pd.notna(v25) and pd.notna(v26):
            delta_bad = float(v26) < float(v25)
        else:
            delta_bad = False

        linhas.append(
            {
                "Linha": linha_ref,
                "2025": v25,
                "2026": v26,
                "2026 Acumulado": v26,
                "2025 Acumulado": v25_acum,
                "Δ R$": delta_rs,
                "Δ %": delta_pct,
                "Alcance 2025": alcance,
                "Ordem": ordem,
                "_delta_bad": delta_bad,
                "_ambos_neg": ambos_neg,
            }
        )
    return pd.DataFrame(linhas)


def formatar_percentual_simples(valor):
    if pd.isna(valor):
        return ""
    try:
        valor = float(valor)
    except Exception:
        return str(valor)
    texto = f"{valor * 100:,.1f}%"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def tabela_html_comparativo(df):
    html = ['<div class="table-wrap"><table class="dash-table">']
    cols = ["Linha", "2025", "2026", "Δ R$", "Δ %", "2025 Acumulado", "Alcance 2025"]
    titulos = {
        "Linha": "Linha",
        "2025": "1Q25",
        "2026": "1Q26",
        "Δ R$": "Δ R$",
        "Δ %": "Δ %",
        "2025 Acumulado": "2025 Acumulado",
        "Alcance 2025": "Alcance 2025",
    }
    html.append("<thead><tr>")
    for col in cols:
        html.append(f"<th>{titulos.get(col, col)}</th>")
    html.append("</tr></thead><tbody>")

    for _, row in df.iterrows():
        row_cls = ' class="total-row"' if normalizar_texto(row["Linha"]) in ["resultado contabil", "resultado contábil"] else ""
        html.append(f"<tr{row_cls}>")
        for col in cols:
            valor = row[col]
            classes = []
            if col in ["2025", "2026", "Δ R$", "2025 Acumulado"]:
                texto = formatar_numero(valor)
                if pd.notna(valor) and valor < 0:
                    classes.append("neg-value")
            elif col in ["Δ %", "Alcance 2025"]:
                texto = formatar_percentual_simples(valor)
                if col == "Δ %" and pd.notna(valor):
                    ambos_neg = row.get("_ambos_neg", False)
                    delta_bad = row.get("_delta_bad", False)
                    valor_exibir = -float(valor) if ambos_neg else float(valor)
                    sinal = "+" if valor_exibir >= 0 else "−"
                    pct = f"{abs(valor_exibir) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")
                    texto = f"{sinal}{pct}"
                    classes.append("delta-negative" if delta_bad else "delta-positive")
            else:
                texto = str(valor)
            cls = f' class="{" ".join(classes)}"' if classes else ""
            html.append(f"<td{cls}>{texto}</td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)


def obter_linha_comparativo(df_comp_principais, linha_ref):
    linha_norm = normalizar_texto(linha_ref)
    return df_comp_principais[df_comp_principais["Linha"].map(normalizar_texto).eq(linha_norm)]


def grafico_alcance_vs_orcado(valor_acumulado, valor_orcado):
    if pd.isna(valor_acumulado) or pd.isna(valor_orcado) or float(valor_orcado) == 0:
        return None

    base = abs(float(valor_orcado))
    realizado = float(valor_acumulado)
    alcance = realizado / base
    alcance_pct = alcance * 100
    eixo_max = max(100.0, alcance_pct * 1.25)
    cor_barra = "#24a8ff" if alcance_pct <= 100 else "#22c55e"

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=alcance_pct,
            number={"suffix": "%", "font": {"size": 62, "color": "#ffffff", "family": "Arial Black"}},
            title={"text": "<b>Resultado Contábil acumulado vs Orçado 2026</b>", "font": {"size": 17, "color": "#ffffff"}},
            gauge={
                "axis": {"range": [0, eixo_max], "tickformat": ".0f", "tickfont": {"color": "#ffffff", "size": 12}},
                "bar": {"color": cor_barra, "thickness": 0.38},
                "bgcolor": "#111a2e",
                "bordercolor": "#243150",
                "borderwidth": 1,
                "steps": [{"range": [0, min(100.0, eixo_max)], "color": "#162338"}],
                "threshold": {"line": {"color": "#ef4444", "width": 4}, "thickness": 0.85, "value": 100},
            },
        )
    )

    diferenca = realizado - base
    if diferenca >= 0:
        texto_status = f"<b>Superou o orçado em:</b> {formatar_moeda(diferenca)}"
    else:
        texto_status = f"<b>Falta para alcançar o orçado:</b> {formatar_moeda(abs(diferenca))}"

    fig.add_annotation(
        x=0.5, y=-0.22,
        xref="paper", yref="paper",
        showarrow=False, align="center",
        text=(
            f"<b>Realizado acumulado:</b> {formatar_moeda(realizado)}<br>"
            f"<b>Orçado 2026:</b> {formatar_moeda(base)}<br>"
            f"{texto_status}"
        ),
        font={"size": 15, "color": "#ffffff"},
    )

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="#080f1f",
        plot_bgcolor="#080f1f",
        height=480,
        margin=dict(l=30, r=30, t=70, b=130),
    )

    return fig


def grafico_alcance_resultado_contabil(valor_2026, valor_base_2025):
    if pd.isna(valor_2026) or pd.isna(valor_base_2025) or float(valor_base_2025) == 0:
        return None

    base = abs(float(valor_base_2025))
    realizado = float(valor_2026)
    alcance = realizado / base
    alcance_pct = alcance * 100
    eixo_max = max(100.0, alcance_pct * 1.25)
    cor_barra = "#24a8ff" if alcance_pct <= 100 else "#22c55e"

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=alcance_pct,
            number={"suffix": "%", "font": {"size": 64, "color": "#ffffff", "family": "Arial Black"}},
            title={"text": "<b>Resultado Contábil 1Q26 x acumulado de 2025</b>", "font": {"size": 22, "color": "#ffffff"}},
            gauge={
                "axis": {"range": [0, eixo_max], "tickformat": ".0f", "tickfont": {"color": "#ffffff", "size": 14}},
                "bar": {"color": cor_barra, "thickness": 0.38},
                "bgcolor": "#111a2e",
                "bordercolor": "#243150",
                "borderwidth": 1,
                "steps": [{"range": [0, min(100.0, eixo_max)], "color": "#162338"}],
                "threshold": {"line": {"color": "#ef4444", "width": 4}, "thickness": 0.85, "value": 100},
            },
        )
    )

    diferenca = realizado - base
    if diferenca >= 0:
        texto_status = f"<b>Superou o acumulado de 2025 em:</b> {formatar_moeda(diferenca)}"
    else:
        texto_status = f"<b>Falta para alcançar o acumulado de 2025:</b> {formatar_moeda(abs(diferenca))}"

    fig.add_annotation(
        x=0.5,
        y=-0.18,
        xref="paper",
        yref="paper",
        showarrow=False,
        align="center",
        text=(
            f"<b>Realizado no 1Q26:</b> {formatar_moeda(realizado)}<br>"
            f"<b>Acumulado de 2025:</b> {formatar_moeda(base)}<br>"
            f"{texto_status}"
        ),
        font={"size": 15, "color": "#ffffff"},
    )

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="#080f1f",
        plot_bgcolor="#080f1f",
        height=580,
        margin=dict(l=80, r=80, t=80, b=140),
    )

    return fig


def montar_tabela_empresas_e_total(df):
    excluir_exatos = {
        "banco",
        "equiv patr",
        "jcp dividendos",
        "br cards",
        "resultado mep",
        "resultado congl financeiro",
        "resultado conglomerado financeiro",
        "resultado coligadas",
        "resultado congl coligadas",
        "resultado conglomerado coligadas",
    }

    excluir_contem = [
        "resultado congl financeiro",
        "resultado conglomerado financeiro",
        "resultado coligadas",
        "resultado congl coligadas",
        "resultado conglomerado coligadas",
    ]

    renomear = {
        "resultado banco": "Banco",
        "resulta br cards": "BR Cards",
        "resultado br cards": "BR Cards",
        "res total": "Resultado Total",
        "resultado total": "Resultado Total",
    }

    linha_total = achar_linha_exata_ou_contendo(df, ["res total", "resultado total"])

    linhas = df[["Linha", "Linha_Normalizada", "Ordem_Linha"]].drop_duplicates().sort_values("Ordem_Linha").copy()

    def manter(row):
        nome = row["Linha_Normalizada"]
        if linha_total is not None and row["Linha"] == linha_total:
            return True
        if nome in excluir_exatos:
            return False
        if any(term in nome for term in excluir_contem):
            return False
        return True

    linhas_filtradas = linhas[linhas.apply(manter, axis=1)]["Linha"].tolist()
    df_tabela = df[df["Linha"].isin(linhas_filtradas)].copy()

    datas_ordem = df_tabela[["Período", "Data"]].drop_duplicates().sort_values("Data")
    colunas = datas_ordem["Período"].tolist()

    tabela = (
        df_tabela.pivot_table(index="Linha", columns="Período", values="Valor", aggfunc="sum")
        .reindex(index=linhas_filtradas)
        .reindex(columns=colunas)
        .reset_index()
    )

    def nome_exibicao(linha):
        nome_norm = normalizar_texto(linha)
        return renomear.get(nome_norm, linha)

    tabela["Linha"] = tabela["Linha"].map(nome_exibicao)
    return tabela


st.sidebar.title("Filtros")
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
        df_2025_acumulado = carregar_2025_acumulado(arquivo)
        df_comp_principais = montar_comparativo_principais(df_comp, df_2025_acumulado)

        if df_comp.empty or df_comp_principais.empty:
            st.info("Não encontrei dados suficientes na aba Comparativo 2026 x 2025.")
        else:
            st.markdown('<div class="section-title">Comparativo 1Q26 x 1Q25</div>', unsafe_allow_html=True)

            # === NOVA CONFIGURAÇÃO DE CARDSTrimestrais SOLICITADOS ===
            # Linha 1: 4 colunas para as primeiras linhas solicitadas
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
                        card(titulo, 0, ajuda="Sem dados na base", variacao=None)
                    else:
                        valor_2025 = linha_df["2025"].iloc[0]
                        valor_2026 = linha_df["2026"].iloc[0]
                        variacao = linha_df["Δ %"].iloc[0]
                        ajuda = f"1Q26: {formatar_moeda(valor_2026)} | 1Q25: {formatar_moeda(valor_2025)}"
                        cor_classe = None
                        variacao_exibir = None
                        if variacao is not None and not pd.isna(variacao) and pd.notna(valor_2026) and float(valor_2026) < 0 and float(variacao) < 0:
                            variacao_exibir = -variacao
                            cor_classe = "delta-negative"
                        card(titulo, valor_2026, ajuda=ajuda, variacao=variacao,
                             variacao_label="Δ 1Q26 vs 1Q25", cor_classe=cor_classe, variacao_exibir=variacao_exibir)

            st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)

            # Linha 2: 3 colunas para as linhas finais de resultados e médias
            novos_cards_linha2 = [
                ("Resultado Contábil 1Q26", "RESULTADO CONTÁBIL"),
                ("Carteira de Crédito (Média do Trimestre) 1Q26", "Carteira de Crédito Bruta Média"),
                ("PL Médio 1Q26", "PL Médio"),
            ]
            
            cols2 = st.columns(3)
            for col, (titulo, linha_nome) in zip(cols2, novos_cards_linha2):
                with col:
                    linha_df = obter_linha_comparativo(df_comp_principais, linha_nome)
                    
                    # Buscas alternativas inteligentes em caso de variações estruturais sutis da planilha
                    if linha_df.empty and "Carteira" in titulo:
                        linha_df = obter_linha_comparativo(df_comp_principais, "Carteira de Crédito Média")
                    if linha_df.empty and "PL" in titulo:
                        linha_df = obter_linha_comparativo(df_comp_principais, "PL Médio (Banco + Hipo)")
                        
                    if linha_df.empty:
                        card(titulo, 0, ajuda="Sem dados na base", variacao=None)
                    else:
                        valor_2025 = linha_df["2025"].iloc[0]
                        valor_2026 = inline_v26 := linha_df["2026"].iloc[0]
                        variacao = linha_df["Δ %"].iloc[0]
                        ajuda = f"1Q26: {formatar_moeda(valor_2026)} | 1Q25: {formatar_moeda(valor_2025)}"
                        cor_classe = None
                        variacao_exibir = None
                        if variacao is not None and not pd.isna(variacao) and pd.notna(valor_2026) and float(valor_2026) < 0 and float(variacao) < 0:
                            variacao_exibir = -variacao
                            cor_classe = "delta-negative"
                        card(titulo, valor_2026, ajuda=ajuda, variacao=variacao,
                             variacao_label="Δ 1Q26 vs 1Q25", cor_classe=cor_classe, variacao_exibir=variacao_exibir)

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
            base_long = df_comp_principais.melt(
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
                margin=dict(l=10, r=120, t=10, b=20),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=13, color="#ffffff", family="Arial Black")),
            )
            st.plotly_chart(fig_comp_ano, use_container_width=True)

            st.markdown('<div class="section-title">Tabela comparativa</div>', unsafe_allow_html=True)
            st.markdown(tabela_html_comparativo(df_comp_principais), unsafe_allow_html=True)

    except Exception as erro:
        st.info(f"Não consegui carregar a aba Comparativo 2026 x 2025: {erro}")
