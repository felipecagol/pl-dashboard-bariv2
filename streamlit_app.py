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


def get_file_hash(arquivo):
    if isinstance(arquivo, Path):
        return arquivo.stat().st_mtime
    if arquivo is not None:
        return arquivo.size
    return 0


def safe_float(valor):
    if pd.isna(valor):
        return pd.NA
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if not texto or texto == '-' or texto.lower() == 'nan':
        return pd.NA
    texto = re.sub(r'[^\d\.,\-]', '', texto)
    if not texto or texto == '-':
        return pd.NA
    if '.' in texto and ',' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    try:
        return float(texto)
    except:
        return pd.NA


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
def carregar_resultado(arquivo, cache_buster=None):
    bruto = pd.read_excel(arquivo, sheet_name=ABA_RESULTADO, header=None, engine="openpyxl")
    bruto = bruto.dropna(how="all").reset_index(drop=True)

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

    for idx in range(linha_mes + 1, len(bruto)):
        linha_nome = bruto.loc[idx, col_rotulo]
        if pd.isna(linha_nome) or str(linha_nome).strip() == "":
            continue

        linha_tem_valor = False
        for col, periodo in colunas_periodo:
            valor = safe_float(bruto.loc[idx, col])
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
def carregar_base_dash(arquivo, cache_buster=None):
    df = pd.read_excel(arquivo, sheet_name=ABA_BASE, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    for col in ["Visao", "Linha_PnL", "Produto", "Metrica", "Periodo"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    if "Valor" in df.columns:
        df["Valor"] = df["Valor"].apply(safe_float).fillna(0)
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

            for c_data in range(col + 1, len(bruto.columns)):
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
def carregar_pnl_mensal(arquivo, cache_buster=None):
    try:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal - Anualizado", header=None, engine="openpyxl")
    except Exception:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Mensal", header=None, engine="openpyxl")

    bruto = bruto.dropna(how="all").reset_index(drop=True)
    registros = []

    for idx in bruto.index:
        for col in bruto.columns:
            if normalizar_texto(bruto.loc[idx, col]) != "data base":
                continue

            linha_data = idx
            col_rotulo = col

            linha_produto = None
            for r_offset in range(1, 6):
                if idx + r_offset < len(bruto):
                    row_vals = [normalizar_texto(str(x)) for x in bruto.iloc[idx + r_offset].values]
                    if "total" in row_vals and ("consignado" in row_vals or "imobiliario" in row_vals):
                        linha_produto = idx + r_offset
                        break
            if not linha_produto:
                continue

            linha_metrica = None
            for r_offset in range(1, 4):
                if linha_produto + r_offset < len(bruto):
                    row_vals = [normalizar_texto(str(x)) for x in bruto.iloc[linha_produto + r_offset].values]
                    if "realizado" in row_vals or "orcado" in row_vals:
                        linha_metrica = linha_produto + r_offset
                        break
            if not linha_metrica:
                continue

            periodos_linha = []
            for c_data in range(col + 1, len(bruto.columns)):
                val = converter_periodo(bruto.loc[linha_data, c_data])
                if val is not None:
                    periodos_linha.append({"col": c_data, "data": val})
            if not periodos_linha:
                continue

            for i_p, p_info in enumerate(periodos_linha):
                data_base = p_info["data"]
                c_inicio_data = p_info["col"]
                c_fim_data = periodos_linha[i_p + 1]["col"] if i_p + 1 < len(periodos_linha) else len(bruto.columns)

                produtos_encontrados = {}
                for c_prod in range(c_inicio_data, c_fim_data):
                    if c_prod not in bruto.columns:
                        continue
                    produto_norm = normalizar_texto(bruto.loc[linha_produto, c_prod])
                    if produto_norm in ["consignado", "imobiliario", "total", "banco digital"]:
                        nome_real = {"consignado": "Consignado", "imobiliario": "Imobiliário", "total": "Total", "banco digital": "Banco Digital"}[produto_norm]
                        if nome_real not in produtos_encontrados:
                            produtos_encontrados[nome_real] = c_prod

                if not {"Consignado", "Imobiliário", "Total"}.issubset(set(produtos_encontrados.keys())):
                    continue

                produtos_ordenados = sorted(produtos_encontrados.items(), key=lambda x: x[1])
                blocos = []

                for i_prod, (produto, col_inicio) in enumerate(produtos_ordenados):
                    col_fim = produtos_ordenados[i_prod + 1][1] if i_prod + 1 < len(produtos_ordenados) else c_fim_data
                    for c_met in range(col_inicio, col_fim):
                        if c_met not in bruto.columns:
                            continue
                        metrica_original = bruto.loc[linha_metrica, c_met]
                        metrica_norm = normalizar_texto(metrica_original)

                        if metrica_norm == "realizado": metrica = "Realizado"
                        elif metrica_norm == "orcado": metrica = "Orçado"
                        elif "r" in metrica_norm and ("r" == metrica_norm or "rs" in metrica_norm): metrica = "Δ R$"
                        elif "%" in str(metrica_original) or "delta" in metrica_norm or metrica_norm in [""]: metrica = "Δ %"
                        else: continue

                        blocos.append({"Produto": "Total" if produto == "Total" else produto, "Coluna": c_met, "Métrica": metrica})

                ordem = 0
                for r in range(linha_metrica + 1, len(bruto)):
                    linha_nome = bruto.loc[r, col_rotulo] if col_rotulo in bruto.columns else None
                    if pd.isna(linha_nome) or str(linha_nome).strip() == "":
                        continue
                    linha_tem_valor = False
                    for bloco in blocos:
                        c_val = bloco["Coluna"]
                        if c_val not in bruto.columns:
                            continue
                        valor = safe_float(bruto.loc[r, c_val])
                        if pd.notna(valor):
                            linha_tem_valor = True
                            registros.append({
                                "Periodo": nome_periodo(data_base),
                                "Data": pd.Timestamp(data_base),
                                "Produto": bloco["Produto"],
                                "Linha": str(linha_nome).strip(),
                                "Linha_Normalizada": normalizar_texto(linha_nome),
                                "Métrica": bloco["Métrica"],
                                "Valor": float(valor),
                                "Ordem_Linha": ordem,
                            })
                    if linha_tem_valor:
                        ordem += 1

    df = pd.DataFrame(registros)
    if df.empty:
        raise ValueError("A aba P&L Mensal - Anualizado foi encontrada, mas nenhum valor numérico foi lido.")

    df = df.drop_duplicates(subset=["Periodo", "Produto", "Linha_Normalizada", "Métrica"], keep="first").reset_index(drop=True)

    linhas_em_modulo = [normalizar_texto("Alíquota de IR/CSLL"), normalizar_texto("Rácio de Eficiência")]
    mask_modulo = (df["Linha_Normalizada"].isin(linhas_em_modulo) & df["Métrica"].isin(["Realizado", "Orçado"]))
    df.loc[mask_modulo, "Valor"] = df.loc[mask_modulo, "Valor"].abs()

    return df


@st.cache_data(show_spinner=False)
def carregar_pnl_acumulado_oficial_completo(arquivo, cache_buster=None):
    try:
        bruto = pd.read_excel(arquivo, sheet_name="P&L Acumulado", header=None, engine="openpyxl")
    except Exception:
        return pd.DataFrame()

    bruto = bruto.dropna(how="all").reset_index(drop=True)
    if bruto.empty:
        return pd.DataFrame()

    label_col = 0
    for c in range(min(5, len(bruto.columns))):
        for r in range(min(15, len(bruto))):
            if normalizar_texto(bruto.iat[r, c]) in ["receitas", "resultado contabil", "resultado contábil"]:
                label_col = c
                break
        if label_col != 0:
            break

    data_ate = None
    for r in range(min(10, len(bruto))):
        for c in range(label_col + 1, min(label_col + 15, len(bruto.columns))):
            val = converter_periodo(bruto.iat[r, c])
            if val is not None:
                data_ate = val
                break
        if data_ate is not None:
            break

    periodo_label = nome_periodo(data_ate) if data_ate is not None else "Acumulado"

    linha_produto = 1
    produtos_map = {}
    for r in range(min(15, len(bruto))):
        row_vals = [normalizar_texto(str(x)) for x in bruto.iloc[r].values]
        if "total" in row_vals and ("consignado" in row_vals or "imobiliario" in row_vals):
            linha_produto = r
            break

    for c in range(label_col + 1, min(label_col + 30, len(bruto.columns))):
        val = normalizar_texto(bruto.iat[linha_produto, c])
        if val in ["consignado", "imobiliario", "total", "banco digital"]:
            nome_real = {"consignado": "Consignado", "imobiliario": "Imobiliário", "total": "Total", "banco digital": "Banco Digital"}[val]
            if nome_real not in produtos_map:
                produtos_map[nome_real] = c

    linha_metrica = linha_produto + 1
    for r in range(linha_produto + 1, min(linha_produto + 4, len(bruto))):
        row_vals = [normalizar_texto(str(x)) for x in bruto.iloc[r].values]
        if "realizado" in row_vals or "orcado" in row_vals:
            linha_metrica = r
            break

    def get_metrics(start_col):
        real_c = start_col
        orc_c = start_col + 1
        for c in range(start_col, min(start_col + 4, len(bruto.columns))):
            m = normalizar_texto(bruto.iat[linha_metrica, c])
            if m == "realizado": real_c = c
            elif m == "orcado": orc_c = c
        return real_c, orc_c

    cols_produto = {}
    for p, c in produtos_map.items():
        r_c, o_c = get_metrics(c)
        cols_produto[p] = {"Realizado": r_c, "Orçado": o_c}

    if not cols_produto:
        cols_produto = {
            "Consignado": {"Realizado": label_col + 1, "Orçado": label_col + 2},
            "Imobiliário": {"Realizado": label_col + 4, "Orçado": label_col + 5},
            "Total": {"Realizado": label_col + 7, "Orçado": label_col + 8},
        }

    registros = []
    ordem = 0
    for r in range(linha_metrica + 1, len(bruto)):
        linha_nome = bruto.iat[r, label_col]
        if pd.isna(linha_nome) or str(linha_nome).strip() == "":
            continue

        linha_tem_valor = False
        linha_norm = normalizar_texto(linha_nome)

        for produto, cols in cols_produto.items():
            col_real = cols.get("Realizado")
            col_orc = cols.get("Orçado")

            val_real = safe_float(bruto.iat[r, col_real]) if col_real and col_real < len(bruto.columns) else pd.NA
            val_orc = safe_float(bruto.iat[r, col_orc]) if col_orc and col_orc < len(bruto.columns) else pd.NA

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

    df = df.drop_duplicates(subset=["Periodo", "Produto", "Linha_Normalizada", "Métrica"], keep="first").reset_index(drop=True)

    linhas_em_modulo = [normalizar_texto("Alíquota de IR/CSLL"), normalizar_texto("Rácio de Eficiência")]
    mask_modulo = (df["Linha_Normalizada"].isin(linhas_em_modulo) & df["Métrica"].isin(["Realizado", "Orçado"]))
    df.loc[mask_modulo, "Valor"] = df.loc[mask_modulo, "Valor"].abs()

    return df


def linhas_ocultas_pnl():
    return {
        normalizar_texto("Componente Juros"), normalizar_texto("Componente Inflação"),
        normalizar_texto("RESULTADO ANTES IMPOSTO RECORRENTE"), normalizar_texto("Rácio de Eficiência Recorrente"),
        normalizar_texto("Alocação de Capital"), normalizar_texto("PL Médio (Banco + Hipo)"),
        normalizar_texto("PL Médio (Prudencial + BRCards)"), normalizar_texto("Carteira de Crédito Bruta Média"),
        normalizar_texto("Carteira de Crédito SD Cliente Média"), normalizar_texto("Taxa Média Carteira Bruta Média"),
        normalizar_texto("Taxa Média Carteira SD Cliente Média"), normalizar_texto("Rateio Carteira"),
        normalizar_texto("Despesas Administrativas")
    }


def obter_linhas_tabela_pnl(df_pnl):
    if df_pnl.empty: return []
    linhas = df_pnl[["Linha", "Linha_Normalizada", "Ordem_Linha"]].drop_duplicates().sort_values("Ordem_Linha").drop_duplicates(subset=["Linha_Normalizada"], keep="first")
    ocultas = linhas_ocultas_pnl()
    linhas = linhas[~linhas["Linha_Normalizada"].isin(ocultas)]
    return linhas["Linha"].tolist()


def obter_linhas_principais_pnl(df_pnl):
    linhas_desejadas = [
        "RECEITAS", "DESPESAS DE ORIGINAÇÃO", "MARGEM INTERMEDIAÇÃO", "Provisões", "MG INTERMEDIAÇÃO LIQ",
        "Despesas Administrativas", "Impostos Diretos", "Comissão (Apoio - Comercial - Mídia Paga)",
        "Amortização ", "RESULTADO ANTES IMPOSTO", "Impostos (IR/CSLL)", "RESULTADO CONTÁBIL"
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
    if df_pnl.empty: return df_pnl
    componentes = [normalizar_texto("Despesas Administrativas Diretas"), normalizar_texto("Desp. Administrativas Indiretas")]
    base_componentes = df_pnl[df_pnl["Linha_Normalizada"].isin(componentes)].copy()
    if base_componentes.empty:
        mask = pd.Series(False, index=df_pnl.index)
        for comp in componentes:
            mask = mask | df_pnl["Linha_Normalizada"].str.contains(comp, regex=False, na=False)
        base_componentes = df_pnl[mask].copy()

    if base_componentes.empty: return df_pnl
    ordem_alvo = df_pnl[df_pnl["Linha_Normalizada"].isin([normalizar_texto("RESULTADO ANTES IMPOSTO")])]["Ordem_Linha"]
    nova_ordem = (ordem_alvo.min() - 0.5) if not ordem_alvo.empty else 99

    cols_grupo = [c for c in ["Produto", "Métrica", "Periodo", "Data"] if c in df_pnl.columns]
    if not cols_grupo: return df_pnl

    agregado_base = base_componentes.groupby(cols_grupo, as_index=False)["Valor"].sum()
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
    nomes_substituir = {normalizar_texto("Despesas Administrativas"), normalizar_texto("Despesas Adm. Diretas + Indiretas")}
    df_sem = df_pnl[~df_pnl["Linha_Normalizada"].isin(nomes_substituir)].copy()
    return pd.concat([df_sem, agregado], ignore_index=True)


def valor_pnl(df, produto, linha, metrica):
    base = df[(df["Produto"] == produto) & (df["Linha"] == linha) & (df["Métrica"] == metrica)]
    if base.empty: return 0
    if "Ordem_Linha" in base.columns: base = base.sort_values("Ordem_Linha")
    return base["Valor"].iloc[0]


def variacao_pnl_mes_anterior(df_pnl_completo, produto, linha, periodo_atual):
    linha_atual = df_pnl_completo[(df_pnl_completo["Produto"] == produto) & (df_pnl_completo["Linha"] == linha) & (df_pnl_completo["Métrica"] == "Realizado") & (df_pnl_completo["Periodo"] == periodo_atual)]
    if linha_atual.empty: return None
    data_atual = linha_atual["Data"].iloc[0]
    anteriores = df_pnl_completo[(df_pnl_completo["Produto"] == produto) & (df_pnl_completo["Linha"] == linha) & (df_pnl_completo["Métrica"] == "Realizado") & (df_pnl_completo["Data"] < data_atual)].sort_values("Data")
    if anteriores.empty: return None
    periodo_anterior = anteriores["Periodo"].iloc[-1]
    valor_atual = valor_pnl(df_pnl_completo[df_pnl_completo["Periodo"] == periodo_atual], produto, linha, "Realizado")
    valor_anterior = valor_pnl(df_pnl_completo[df_pnl_completo["Periodo"] == periodo_anterior], produto, linha, "Realizado")
    if valor_anterior == 0: return None
    return (valor_atual - valor_anterior) / abs(valor_anterior)


def filtrar_pnl_acumulado(df_pnl_completo, periodo_atual):
    linha_periodo = df_pnl_completo[df_pnl_completo["Periodo"] == periodo_atual]
    if linha_periodo.empty: return df_pnl_completo.iloc[0:0].copy()
    data_atual = linha_periodo["Data"].iloc[0]
    ano_atual = pd.Timestamp(data_atual).year
    data_inicio = pd.Timestamp(ano_atual, 1, 1)
    base = df_pnl_completo[(df_pnl_completo["Data"] >= data_inicio) & (df_pnl_completo["Data"] <= data_atual)].copy()
    return base


def buscar_linha_acumulada(df_acumulado, produto, linha_norm_alvo, metrica="Realizado"):
    base = df_acumulado[(df_acumulado["Produto"] == produto) & (df_acumulado["Linha_Normalizada"] == linha_norm_alvo) & (df_acumulado["Métrica"] == metrica)]
    if base.empty: return None
    return float(base["Valor"].iloc[0])


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


def recalcular_indicadores_percentuais(df_acumulado, n_meses):
    if df_acumulado.empty or n_meses <= 0: return df_acumulado
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

        def calc_margem_bruta(mg, cart): return mg * fator_anual / cart if (mg is not None and cart and cart != 0) else None
        def calc_margem_liquida(mg_liq, cart): return mg_liq * fator_anual / cart if (mg_liq is not None and cart and cart != 0) else None
        def calc_racio_eficiencia(d_dir, d_ind, mg): return abs((d_dir + d_ind) / mg) if (d_dir is not None and d_ind is not None and mg and mg != 0) else None
        def calc_rpl(res, aloc): return res * fator_anual / aloc if (res is not None and aloc and aloc != 0) else None
        def calc_aliquota(imp, rai_v): return abs(imp / rai_v) if (imp is not None and rai_v and rai_v != 0) else None

        indicadores = [
            ("Margem Bruta", normalizar_texto("Margem Bruta"), calc_margem_bruta(mg_int, carteira_bruta), calc_margem_bruta(mg_int_orc, carteira_bruta_orc)),
            ("Margem Liquida", normalizar_texto("Margem Liquida"), calc_margem_liquida(mg_int_liq, carteira_bruta), calc_margem_liquida(mg_int_liq_orc, carteira_bruta_orc)),
            ("Rácio de Eficiência", normalizar_texto("Rácio de Eficiência"), calc_racio_eficiencia(desp_adm_dir, desp_adm_ind, mg_int), calc_racio_eficiencia(desp_adm_dir_orc, desp_adm_ind_orc, mg_int_orc)),
            ("RPL - RES. CONTÁBIL", normalizar_texto("RPL - RES. CONTÁBIL"), calc_rpl(res_contabil, alocacao_capital), calc_rpl(res_contabil_orc, alocacao_capital_orc)),
            ("Alíquota de IR/CSLL", normalizar_texto("Alíquota de IR/CSLL"), calc_aliquota(impostos, rai), calc_aliquota(impostos_orc, rai_orc)),
        ]

        for nome_exibir, linha_norm, val_real, val_orc in indicadores:
            ordem_existente = df_acumulado[df_acumulado["Linha_Normalizada"] == linha_norm]["Ordem_Linha"]
            ordem = ordem_existente.iloc[0] if not ordem_existente.empty else 9999
            if val_real is not None:
                novos_registros.append({"Produto": produto, "Linha": nome_exibir, "Linha_Normalizada": linha_norm, "Métrica": "Realizado", "Ordem_Linha": ordem, "Valor": val_real})
            if val_orc is not None:
                novos_registros.append({"Produto": produto, "Linha": nome_exibir, "Linha_Normalizada": linha_norm, "Métrica": "Orçado", "Ordem_Linha": ordem, "Valor": val_orc})

    if not novos_registros: return df_acumulado
    linhas_recalculadas = {r["Linha_Normalizada"] for r in novos_registros}
    df_sem_percentuais = df_acumulado[~(df_acumulado["Linha_Normalizada"].isin(linhas_recalculadas) & df_acumulado["Métrica"].isin(["Realizado", "Orçado"]))].copy()
    return pd.concat([df_sem_percentuais, pd.DataFrame(novos_registros)], ignore_index=True)


def agregar_pnl_acumulado(df_pnl_periodo):
    if df_pnl_periodo.empty: return df_pnl_periodo.copy()
    linhas_percentuais = linhas_percentuais_pnl()
    linhas_media = linhas_media_pnl()
    base_valores = df_pnl_periodo[df_pnl_periodo["Métrica"].isin(["Realizado", "Orçado"])].copy()
    n_meses = base_valores["Periodo"].nunique() if "Periodo" in base_valores.columns else 1
    if n_meses <= 0: n_meses = 1

    mask_pct = base_valores["Linha_Normalizada"].isin(linhas_percentuais)
    mask_media = base_valores["Linha_Normalizada"].isin(linhas_media)
    mask_soma = ~(mask_pct | mask_media)
    
    agrupado_soma = base_valores[mask_soma].groupby(["Produto", "Linha", "Linha_Normalizada", "Métrica"], as_index=False)["Valor"].sum()
    if not base_valores[mask_media].empty:
        agrupado_media = base_valores[mask_media].groupby(["Produto", "Linha", "Linha_Normalizada", "Métrica"], as_index=False)["Valor"].mean()
    else:
        agrupado_media = base_valores[mask_media].iloc[0:0][["Produto", "Linha", "Linha_Normalizada", "Métrica", "Valor"]]
        
    agrupado = pd.concat([agrupado_soma, agrupado_media], ignore_index=True)
    ordem_por_linha = base_valores.groupby("Linha_Normalizada", as_index=False)["Ordem_Linha"].min()
    agrupado = agrupado.merge(ordem_por_linha, on="Linha_Normalizada", how="left")
    agrupado = recalcular_indicadores_percentuais(agrupado, n_meses)

    linhas_delta = []
    base_pivot = agrupado.pivot_table(index=["Produto", "Linha", "Linha_Normalizada", "Ordem_Linha"], columns="Métrica", values="Valor", aggfunc="first").reset_index()
    for _, row in base_pivot.iterrows():
        realizado = row.get("Realizado")
        orcado = row.get("Orçado")
        linha_norm = row["Linha_Normalizada"]
        eh_percentual = linha_norm in linhas_percentuais

        if pd.isna(realizado) or pd.isna(orcado):
            delta_rs = pd.NA
            delta_pct = pd.NA
        else:
            delta_rs = realizado - orcado
            if eh_percentual: delta_pct = delta_rs
            else: delta_pct = pd.NA if orcado == 0 else delta_rs / abs(orcado)

        for metrica, valor in [("Δ %", delta_pct), ("Δ R$", delta_rs)]:
            linhas_delta.append({"Produto": row["Produto"], "Linha": row["Linha"], "Linha_Normalizada": row["Linha_Normalizada"], "Métrica": metrica, "Ordem_Linha": row["Ordem_Linha"], "Valor": valor})
            
    return pd.concat([agrupado, pd.DataFrame(linhas_delta)], ignore_index=True)


def variacao_pnl_acumulado_mes_anterior(df_pnl_completo, produto, linha, periodo_atual):
    linha_atual = df_pnl_completo[(df_pnl_completo["Produto"] == produto) & (df_pnl_completo["Linha"] == linha) & (df_pnl_completo["Métrica"] == "Realizado") & (df_pnl_completo["Periodo"] == periodo_atual)]
    if linha_atual.empty: return None
    data_atual = linha_atual["Data"].iloc[0]
    ano_atual = pd.Timestamp(data_atual).year
    data_inicio = pd.Timestamp(ano_atual, 1, 1)

    meses_anteriores = df_pnl_completo[(df_pnl_completo["Produto"] == produto) & (df_pnl_completo["Linha"] == linha) & (df_pnl_completo["Métrica"] == "Realizado") & (df_pnl_completo["Data"] < data_atual) & (df_pnl_completo["Data"] >= data_inicio)].sort_values("Data")
    if meses_anteriores.empty: return None

    valor_acumulado_atual = df_pnl_completo[(df_pnl_completo["Produto"] == produto) & (df_pnl_completo["Linha"] == linha) & (df_pnl_completo["Métrica"] == "Realizado") & (df_pnl_completo["Data"] >= data_inicio) & (df_pnl_completo["Data"] <= data_atual)]["Valor"].sum()
    data_anterior = meses_anteriores["Data"].max()
    valor_acumulado_anterior = df_pnl_completo[(df_pnl_completo["Produto"] == produto) & (df_pnl_completo["Linha"] == linha) & (df_pnl_completo["Métrica"] == "Realizado") & (df_pnl_completo["Data"] >= data_inicio) & (df_pnl_completo["Data"] <= data_anterior)]["Valor"].sum()

    if valor_acumulado_anterior == 0: return None
    return (valor_acumulado_atual - valor_acumulado_anterior) / abs(valor_acumulado_anterior)


def variacao_pnl_acumulado_vs_2025(df_comp_2025, produto, linha, valor_ytd_atual):
    if df_comp_2025 is None or df_comp_2025.empty or pd.isna(valor_ytd_atual): return None
    linha_norm = normalizar_texto(linha)
    NORM_DESP_ADM = normalizar_texto("Despesas Administrativas")
    NORM_DESP_DIR = normalizar_texto("Despesas Administrativas Diretas")
    NORM_DESP_IND = normalizar_texto("Desp. Administrativas Indiretas")

    def buscar_valor_2025(norm_alvo):
        if produto == "Total":
            if "comissao" in norm_alvo: return -2314689.0
            if "amortizacao" in norm_alvo: return -3247662.0
        base = df_comp_2025[(df_comp_2025["Ano"] == 2025) & (df_comp_2025["Produto"] == produto) & (df_comp_2025["Linha_Normalizada"] == norm_alvo)]
        if base.empty:
            if "comissao" in norm_alvo: base = df_comp_2025[(df_comp_2025["Ano"] == 2025) & (df_comp_2025["Produto"] == produto) & (df_comp_2025["Linha_Normalizada"].str.contains("comissao", na=False))]
            elif "amortizacao" in norm_alvo: base = df_comp_2025[(df_comp_2025["Ano"] == 2025) & (df_comp_2025["Produto"] == produto) & (df_comp_2025["Linha_Normalizada"].str.contains("amortizacao", na=False))]
            elif "impostos diretos" in norm_alvo: base = df_comp_2025[(df_comp_2025["Ano"] == 2025) & (df_comp_2025["Produto"] == produto) & (df_comp_2025["Linha_Normalizada"].str.contains("impostos diretos", na=False))]
        if base.empty: return None
        v = pd.to_numeric(base["Realizado"].iloc[0], errors="coerce")
        return float(v) if pd.notna(v) else None

    if linha_norm == NORM_DESP_ADM:
        v_dir = buscar_valor_2025(NORM_DESP_DIR)
        v_ind = buscar_valor_2025(NORM_DESP_IND)
        if v_dir is None and v_ind is None: return None
        valor_2025 = (v_dir or 0.0) + (v_ind or 0.0)
    else:
        valor_2025 = buscar_valor_2025(linha_norm)

    if valor_2025 is None or valor_2025 == 0: return None
    return (float(valor_ytd_atual) - valor_2025) / abs(valor_2025)


def tabela_html_comparativo(df):
    html = ['<div class="table-wrap"><table class="dash-table">']
    cols = ["Linha", "2025", "2026", "Δ R$", "Δ %", "2025 Acumulado", "Alcance 2025"]
    titulos = {"Linha": "Linha", "2025": "1º Quad. 25", "2026": "1º Quad. 26", "Δ R$": "Δ R$", "Δ %": "Δ %", "2025 Acumulado": "2025 Acumulado", "Alcance 2025": "Alcance 2025"}
    html.append("<thead><tr>")
    for col in cols: html.append(f"<th>{titulos.get(col, col)}</th>")
    html.append("</tr></thead><tbody>")
    for _, row in df.iterrows():
        row_cls = ' class="total-row"' if normalizar_texto(row["Linha"]) in ["resultado contabil", "resultado contábil"] else ""
        html.append(f"<tr{row_cls}>")
        for col in cols:
            valor = row[col]
            classes = []
            if col in ["2025", "2026", "Δ R$", "2025 Acumulado"]:
                texto = formatar_numero(valor)
                if pd.notna(valor) and valor < 0: classes.append("neg-value")
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
    res = df_comp_principais[df_comp_principais["Linha"].map(normalizar_texto).eq(linha_norm)]
    if not res.empty: return res
    if "carteira" in linha_norm: res = df_comp_principais[df_comp_principais["Linha"].map(normalizar_texto).str.contains("carteira", na=False)]
    if res.empty and "pl medio" in linha_norm: res = df_comp_principais[df_comp_principais["Linha"].map(normalizar_texto).str.contains("pl medio", na=False)]
    return res


def grafico_alcance_vs_orcado(valor_acumulado, valor_orcado):
    if pd.isna(valor_acumulado) or pd.isna(valor_orcado) or float(valor_orcado) == 0: return None
    base = abs(float(valor_orcado))
    realizado = float(valor_acumulado)
    alcance = realizado / base
    alcance_pct = alcance * 100
    eixo_max = max(100.0, alcance_pct * 1.25)
    cor_barra = "#24a8ff" if alcance_pct <= 100 else "#22c55e"
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=alcance_pct, number={"suffix": "%", "font": {"size": 62, "color": "#ffffff", "family": "Arial Black"}},
        title={"text": "<b>Resultado Contábil acumulado vs Orçado 2026</b>", "font": {"size": 17, "color": "#ffffff"}},
        gauge={
            "axis": {"range": [0, eixo_max], "tickformat": ".0f", "tickfont": {"color": "#ffffff", "size": 12}},
            "bar": {"color": cor_barra, "thickness": 0.38},
            "bgcolor": "#111a2e", "bordercolor": "#243150", "borderwidth": 1,
            "steps": [{"range": [0, min(100.0, eixo_max)], "color": "#162338"}],
            "threshold": {"line": {"color": "#ef4444", "width": 4}, "thickness": 0.85, "value": 100},
        }))
    diferenca = realizado - base
    texto_status = f"<b>Superou o orçado em:</b> {formatar_moeda(diferenca)}" if diferenca >= 0 else f"<b>Falta para alcançar o orçado:</b> {formatar_moeda(abs(diferenca))}"
    fig.add_annotation(
        x=0.5, y=-0.22, xref="paper", yref="paper", showarrow=False, align="center",
        text=(f"<b>Realizado acumulado:</b> {formatar_moeda(realizado)}<br><b>Orçado 2026:</b> {formatar_moeda(base)}<br>{texto_status}"),
        font={"size": 15, "color": "#ffffff"},
    )
    fig.update_layout(template="plotly_dark", paper_bgcolor="#080f1f", plot_bgcolor="#080f1f", height=480, margin=dict(l=30, r=30, t=70, b=130))
    return fig


def grafico_alcance_resultado_contabil(valor_2026, valor_base_2025):
    if pd.isna(valor_2026) or pd.isna(valor_base_2025) or float(valor_base_2025) == 0: return None
    base = abs(float(valor_base_2025))
    realizado = float(valor_2026)
    alcance = realizado / base
    alcance_pct = alcance * 100
    eixo_max = max(100.0, alcance_pct * 1.25)
    cor_barra = "#24a8ff" if alcance_pct <= 100 else "#22c55e"
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=alcance_pct, number={"suffix": "%", "font": {"size": 64, "color": "#ffffff", "family": "Arial Black"}},
        title={"text": "<b>Resultado Contábil 1º Quad. 26 x acumulado de 2025</b>", "font": {"size": 22, "color": "#ffffff"}},
        gauge={
            "axis": {"range": [0, eixo_max], "tickformat": ".0f", "tickfont": {"color": "#ffffff", "size": 14}},
            "bar": {"color": cor_barra, "thickness": 0.38},
            "bgcolor": "#111a2e", "bordercolor": "#243150", "borderwidth": 1,
            "steps": [{"range": [0, min(100.0, eixo_max)], "color": "#162338"}],
            "threshold": {"line": {"color": "#ef4444", "width": 4}, "thickness": 0.85, "value": 100},
        }))
    diferenca = realizado - base
    texto_status = f"<b>Superou o acumulado de 2025 em:</b> {formatar_moeda(diferenca)}" if diferenca >= 0 else f"<b>Falta para alcançar o acumulado de 2025:</b> {formatar_moeda(abs(diferenca))}"
    fig.add_annotation(
        x=0.5, y=-0.18, xref="paper", yref="paper", showarrow=False, align="center",
        text=(f"<b>Realizado no 1º Quad. 26:</b> {formatar_moeda(realizado)}<br><b>Acumulado de 2025:</b> {formatar_moeda(base)}<br>{texto_status}"),
        font={"size": 15, "color": "#ffffff"},
    )
    fig.update_layout(template="plotly_dark", paper_bgcolor="#080f1f", plot_bgcolor="#080f1f", height=580, margin=dict(l=80, r=80, t=80, b=140))
    return fig


# ==============================================================================
# INÍCIO DA EXECUÇÃO PRINCIPAL
# ==============================================================================

st.sidebar.title("Configurações")
arquivo_local = Path(ARQUIVO_PADRAO)
upload = st.sidebar.file_uploader("Atualizar base", type=["xlsx"])
arquivo = upload if upload else (arquivo_local if arquivo_local.exists() else None)

if arquivo:
    file_hash = get_file_hash(arquivo)
    try:
        df_resultado = carregar_resultado(arquivo, file_hash)
    except Exception as erro:
        st.error(f"Erro ao carregar a aba RESULTADO: {erro}")
        st.stop()

    try:
        df_pnl_completo_global = carregar_pnl_mensal(arquivo, file_hash)
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
                xaxis_showgrid=False,
                yaxis_showgrid=False,
                xaxis_zeroline=False,
                yaxis_zeroline=False
            )
            fig.update_xaxes(
                tickmode="array",
                tickvals=tick_datas,
                ticktext=tick_textos,
                range=[x_min, x_max],
                title_text="",
                showgrid=False,
                zeroline=False,
                showline=False,
                tickfont=dict(color='#FFFFFF', size=13, family="Arial", weight="bold")
            )
            fig.update_yaxes(
                tickprefix="R$ ",
                separatethousands=True,
                range=[y_min - y_pad, y_max + y_pad],
                title_text="",
                showgrid=False,
                zeroline=False,
                showline=False
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
            periodos_pnl_mensal = obter_periodos_pnl_mensal_anualizado(arquivo)
            lista_periodos_pnl_mensal = [item["Período"] for item in periodos_pnl_mensal]
            label_para_periodo_mensal = None
            
            st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
            col_data, col_produto, col_espaco = st.columns([1, 1, 2.5])
            with col_data:
                data_sel_pnl = st.selectbox("Data base", lista_periodos_pnl_mensal, index=len(lista_periodos_pnl_mensal) - 1, key="data_pnl_mensal")
            with col_produto:
                produto_sel_pnl = st.selectbox("Produto", ["Consignado", "Imobiliário", "Banco Digital", "Total"], index=3, key="produto_pnl_mensal")

            df_pnl = df_pnl_completo_global[df_pnl_completo_global["Periodo"] == data_sel_pnl].copy()
            linhas_principais = obter_linhas_principais_pnl(df_pnl)

            st.markdown('<div class="section-title">Principais linhas do P&L Mensal</div>', unsafe_allow_html=True)
            ordem_exibicao_cards = ["RECEITAS", "DESPESAS DE ORIGINAÇÃO", "MARGEM INTERMEDIAÇÃO", "Provisões", "MG INTERMEDIAÇÃO LIQ", "Despesas Administrativas", "Impostos Diretos", "Comissão (Apoio - Comercial - Mídia Paga)", "Amortização ", "RESULTADO ANTES IMPOSTO", "Impostos (IR/CSLL)", "RESULTADO CONTÁBIL"]
            linhas_mapeadas = []
            for ref_nome in ordem_exibicao_cards:
                for l in linhas_principais:
                    if normalizar_texto(ref_nome) in normalizar_texto(l):
                        if l not in linhas_mapeadas: linhas_mapeadas.append(l)
                        break
            for l in linhas_principais:
                if l not in linhas_mapeadas:
                    linhas_mapeadas.append(l)

            for inicio in range(0, len(linhas_mapeadas), 3):
                if inicio > 0: st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)
                cols_cards = st.columns(3)
                for col_card, linha in zip(cols_cards, linhas_mapeadas[inicio:inicio + 3]):
                    realizado = valor_pnl(df_pnl, produto_sel_pnl, linha, "Realizado")
                    variacao = variacao_pnl_mes_anterior(df_pnl_completo_global, produto_sel_pnl, linha, data_sel_pnl)
                    cor_classe = "delta-negative" if variacao is not None and not pd.isna(variacao) and realizado < 0 and float(variacao) < 0 else "delta-positive" if variacao is not None and not pd.isna(variacao) and realizado < 0 else None
                    variacao_exibir = -variacao if variacao is not None and not pd.isna(variacao) and realizado < 0 else variacao
                    with col_card: card_pnl(linha.replace("*","").strip(), realizado, variacao=variacao, variacao_label="Δ mês anterior", cor_classe=cor_classe, variacao_exibir=variacao_exibir)

            st.markdown('<div class="section-title">Realizado x Orçado por linha principal</div>', unsafe_allow_html=True)
            novas_medidas = ["impostos diretos", "comissao", "amortizacao"]
            linhas_grafico_pnl = [l for l in linhas_mapeadas if not any(m in normalizar_texto(l) for m in novas_medidas)]
            base_grafico = df_pnl[(df_pnl["Produto"] == produto_sel_pnl) & (df_pnl["Linha"].isin(linhas_grafico_pnl)) & (df_pnl["Métrica"].isin(["Realizado", "Orçado"]))].copy()
            ordem_linhas = {linha: i for i, linha in enumerate(linhas_grafico_pnl)}
            base_grafico["Ordem"] = base_grafico["Linha"].map(ordem_linhas)
            base_grafico = base_grafico.sort_values("Ordem", ascending=False)
            base_grafico["Rótulo"] = base_grafico["Valor"].map(formatar_moeda_curta)
            
            fig_comp = px.bar(base_grafico, x="Valor", y="Linha", color="Métrica", text="Rótulo", orientation="h", barmode="group", labels={"Valor": "Valor", "Linha": "", "Métrica": ""})
            fig_comp.update_traces(texttemplate="<b>%{text}</b>", textposition="outside", textfont=dict(size=14, family="Arial Black", color="#FFFFFF"), cliponaxis=False, constraintext="none")
            fig_comp.update_layout(template="plotly_dark", paper_bgcolor="#080f1f", plot_bgcolor="#080f1f", height=640, margin=dict(l=10, r=140, t=20, b=10), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=13, color="#ffffff", family="Arial Black")), uniformtext_minsize=11, uniformtext_mode="show", bargap=0.30, bargroupgap=0.18, xaxis_showgrid=False, yaxis_showgrid=False, xaxis_zeroline=False, yaxis_zeroline=False)
            if not base_grafico.empty:
                x_min, x_max = base_grafico["Valor"].min(), base_grafico["Valor"].max()
                x_pad = max((x_max - x_min) * 0.32, 1)
                fig_comp.update_xaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickprefix="R$ ", separatethousands=True, range=[x_min - x_pad, x_max + x_pad])
            else: fig_comp.update_xaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickprefix="R$ ", separatethousands=True)
            fig_comp.update_yaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickfont=dict(size=12, color="#ffffff"))
            st.plotly_chart(fig_comp, use_container_width=True)

            st.markdown('<div class="section-title">Resultado Contábil por produto</div>', unsafe_allow_html=True)
            linha_resultado_contabil = next((linha for linha in linhas_mapeadas if normalizar_texto(linha) in ["resultado contabil", "resultado contábil"]), linhas_mapeadas[-1] if len(linhas_mapeadas) > 0 else None)
            base_produtos = df_pnl[(df_pnl["Linha"] == linha_resultado_contabil) & (df_pnl["Produto"].isin(["Consignado", "Imobiliário", "Banco Digital", "Total"])) & (df_pnl["Métrica"] == "Realizado")].copy()
            def texto_barra_mensal(row):
                val_str = formatar_moeda(row["Valor"])
                var = variacao_pnl_mes_anterior(df_pnl_completo_global, row["Produto"], linha_resultado_contabil, data_sel_pnl)
                if var is None or pd.isna(var): return val_str
                sinal = "+" if float(var) >= 0 else "−"
                pct = f"{abs(float(var)) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"{val_str}<br><span style='font-size:14px; font-weight: bold; color: #ffffff;'>{sinal}{pct} vs mês ant.</span>"
            base_produtos["Rótulo"] = base_produtos.apply(texto_barra_mensal, axis=1)
            threshold = (base_produtos["Valor"].max() if not base_produtos.empty else 1) * 0.20
            text_positions = ["outside" if row["Valor"] < threshold else "inside" for _, row in base_produtos.iterrows()]
            fig_prod = go.Figure(go.Bar(x=base_produtos["Produto"], y=base_produtos["Valor"], text=base_produtos["Rótulo"], textposition=text_positions, textfont=dict(size=16, family="Arial Black", color="#FFFFFF"), insidetextanchor="middle", texttemplate="%{text}", marker_color="#1d6ff2"))
            fig_prod.update_layout(template="plotly_dark", paper_bgcolor="#080f1f", plot_bgcolor="#080f1f", height=390, margin=dict(l=10, r=10, t=10, b=10), showlegend=False, xaxis_showgrid=False, yaxis_showgrid=False, xaxis_zeroline=False, yaxis_zeroline=False)
            fig_prod.update_xaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)")
            fig_prod.update_yaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickprefix="R$ ", separatethousands=True)
            st.plotly_chart(fig_prod, use_container_width=True)

            st.markdown('<div class="section-title">Resumo das linhas principais por produto</div>', unsafe_allow_html=True)
            st.markdown(tabela_html_pnl_matriz(montar_matriz_pnl_excel(df_pnl, obter_linhas_tabela_pnl(df_pnl))[0], montar_matriz_pnl_excel(df_pnl, obter_linhas_tabela_pnl(df_pnl))[1], montar_matriz_pnl_excel(df_pnl, obter_linhas_tabela_pnl(df_pnl))[2]), unsafe_allow_html=True)
        else:
            st.info(f"Não consegui carregar a aba P&L Mensal: {erro_pnl_global}")

    with tab_pnl_acum:
        if erro_pnl_global is None and not df_pnl_completo_global.empty:
            try: df_comp_para_acum = carregar_comparativo_2025(arquivo)
            except Exception: df_comp_para_acum = None

            periodos_pnl_acum = obter_periodos_pnl_mensal_anualizado(arquivo)
            meses_fim_trimestre = {3, 4, 5, 6, 9, 12}
            periodos_pnl_acum = [p for p in periodos_pnl_acum if p["Data"] is not None and p["Data"].month in meses_fim_trimestre]
            if not periodos_pnl_acum: periodos_pnl_acum = obter_periodos_pnl_mensal_anualizado(arquivo)
            mes_para_trimestre = {3: "1º Trim.", 4: "Jan-Abr", 5: "Jan-Mai", 6: "1º Sem.", 9: "9M", 12: "Ano"}
            label_para_periodo = {}
            lista_labels_trimestre = []
            for p in periodos_pnl_acum:
                if p["Data"] is not None:
                    tri = mes_para_trimestre.get(p["Data"].month, p["Período"])
                    ano = str(p["Data"].year)[-2:]
                    label = f"{tri} {ano}"
                else: label = p["Período"]
                label_para_periodo[label] = p["Período"]
                lista_labels_trimestre.append(label)

            st.markdown('<div class="section-title">Filtros</div>', unsafe_allow_html=True)
            col_produto_acum, col_espaco_acum = st.columns([1, 3.5])
            df_pnl_acumulado = carregar_pnl_acumulado_oficial_completo(arquivo)
            df_pnl_acumulado = garantir_linha_despesas_administrativas(df_pnl_acumulado)
            data_sel_pnl_acum = df_pnl_acumulado["Periodo"].iloc[0] if not df_pnl_acumulado.empty else "Acumulado"

            with col_produto_acum:
                produto_sel_pnl_acum = st.selectbox("Produto", ["Consignado", "Imobiliário", "Banco Digital", "Total"], index=3, key="produto_pnl_acumulado")

            linhas_principais_acum = obter_linhas_principais_pnl(df_pnl_acumulado)
            st.markdown('<div class="section-title">Principais linhas do P&L Acumulado</div>', unsafe_allow_html=True)

            ordem_exibicao_cards_acum = ["RECEITAS", "DESPESAS DE ORIGINAÇÃO", "MARGEM INTERMEDIAÇÃO", "Provisões", "MG INTERMEDIAÇÃO LIQ", "Despesas Administrativas", "Impostos Diretos", "Comissão (Apoio - Comercial - Mídia Paga)", "Amortização ", "RESULTADO ANTES IMPOSTO", "Impostos (IR/CSLL)", "RESULTADO CONTÁBIL"]
            linhas_mapeadas_acum = []
            for ref_nome in ordem_exibicao_cards_acum:
                for l in linhas_principais_acum:
                    if normalizar_texto(ref_nome) in normalizar_texto(l):
                        if l not in linhas_mapeadas_acum: linhas_mapeadas_acum.append(l)
                        break
            for l in linhas_principais_acum:
                if l not in linhas_mapeadas_acum: linhas_mapeadas_acum.append(l)

            for inicio in range(0, len(linhas_mapeadas_acum), 3):
                if inicio > 0: st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)
                cols_cards = st.columns(3)
                for col_card, linha in zip(cols_cards, linhas_mapeadas_acum[inicio:inicio + 3]):
                    realizado = valor_pnl(df_pnl_acumulado, produto_sel_pnl_acum, linha, "Realizado")
                    variacao = variacao_pnl_acumulado_vs_2025(df_comp_para_acum, produto_sel_pnl_acum, linha, realizado)
                    cor_classe = "delta-negative" if variacao is not None and not pd.isna(variacao) and realizado < 0 and float(variacao) < 0 else "delta-positive" if variacao is not None and not pd.isna(variacao) and realizado < 0 else None
                    variacao_exibir = -variacao if variacao is not None and not pd.isna(variacao) and realizado < 0 else variacao
                    with col_card: card_pnl(linha.replace("*","").strip(), realizado, variacao=variacao, variacao_label="vs Jan-Mai 25", cor_classe=cor_classe, variacao_exibir=variacao_exibir)

            st.markdown('<div class="section-title">Realizado x Orçado acumulado por linha principal</div>', unsafe_allow_html=True)
            novas_medidas_acum = ["impostos diretos", "comissao", "amortizacao"]
            linhas_grafico_pnl_acum = [l for l in linhas_mapeadas_acum if not any(m in normalizar_texto(l) for m in novas_medidas_acum)]
            base_grafico_acum = df_pnl_acumulado[(df_pnl_acumulado["Produto"] == produto_sel_pnl_acum) & (df_pnl_acumulado["Linha"].isin(linhas_grafico_pnl_acum)) & (df_pnl_acumulado["Métrica"].isin(["Realizado", "Orçado"]))].copy()
            ordem_linhas_acum = {linha: i for i, linha in enumerate(linhas_grafico_pnl_acum)}
            base_grafico_acum["Ordem"] = base_grafico_acum["Linha"].map(ordem_linhas_acum)
            base_grafico_acum = base_grafico_acum.sort_values("Ordem", ascending=False)
            base_grafico_acum["Rótulo"] = base_grafico_acum["Valor"].map(formatar_moeda_curta)

            fig_comp_acum = px.bar(base_grafico_acum, x="Valor", y="Linha", color="Métrica", text="Rótulo", orientation="h", barmode="group", labels={"Valor": "Valor", "Linha": "", "Métrica": ""})
            fig_comp_acum.update_traces(texttemplate="<b>%{text}</b>", textposition="outside", textfont=dict(size=14, family="Arial Black", color="#FFFFFF"), cliponaxis=False, constraintext="none")
            fig_comp_acum.update_layout(template="plotly_dark", paper_bgcolor="#080f1f", plot_bgcolor="#080f1f", height=640, margin=dict(l=10, r=140, t=20, b=10), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=13, color="#ffffff", family="Arial Black")), uniformtext_minsize=11, uniformtext_mode="show", bargap=0.30, bargroupgap=0.18, xaxis_showgrid=False, yaxis_showgrid=False, xaxis_zeroline=False, yaxis_zeroline=False)
            if not base_grafico_acum.empty:
                x_min, x_max = base_grafico_acum["Valor"].min(), base_grafico_acum["Valor"].max()
                x_pad = max((x_max - x_min) * 0.32, 1)
                fig_comp_acum.update_xaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickprefix="R$ ", separatethousands=True, range=[x_min - x_pad, x_max + x_pad])
            else: fig_comp_acum.update_xaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickprefix="R$ ", separatethousands=True)
            fig_comp_acum.update_yaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickfont=dict(size=12, color="#ffffff"))
            st.plotly_chart(fig_comp_acum, use_container_width=True)

            st.markdown('<div class="section-title">Resultado Contábil acumulado por produto</div>', unsafe_allow_html=True)
            linha_resultado_contabil_acum = next((linha for linha in linhas_mapeadas_acum if normalizar_texto(linha) in ["resultado contabil", "resultado contábil"]), linhas_mapeadas_acum[-1] if len(linhas_mapeadas_acum) > 0 else None)
            base_produtos_acum = df_pnl_acumulado[(df_pnl_acumulado["Linha"] == linha_resultado_contabil_acum) & (df_pnl_acumulado["Produto"].isin(["Consignado", "Imobiliário", "Banco Digital", "Total"])) & (df_pnl_acumulado["Métrica"] == "Realizado")].copy()
            def texto_barra_acum(row):
                val_str = formatar_moeda(row["Valor"])
                if df_comp_para_acum is not None and not df_comp_para_acum.empty:
                    linha_norm = normalizar_texto(linha_resultado_contabil_acum)
                    b26 = df_comp_para_acum[(df_comp_para_acum["Ano"] == 2026) & (df_comp_para_acum["Produto"] == row["Produto"]) & (df_comp_para_acum["Linha_Normalizada"] == linha_norm)]
                    b25 = df_comp_para_acum[(df_comp_para_acum["Ano"] == 2025) & (df_comp_para_acum["Produto"] == row["Produto"]) & (df_comp_para_acum["Linha_Normalizada"] == linha_norm)]
                    v26 = pd.to_numeric(b26["Realizado"].iloc[0], errors="coerce") if not b26.empty else None
                    v25 = pd.to_numeric(b25["Realizado"].iloc[0], errors="coerce") if not b25.empty else None
                    var = (float(v26) / float(v25)) - 1 if v26 is not None and v25 is not None and pd.notna(v26) and pd.notna(v25) and v25 != 0 else None
                else: var = None
                if var is None or pd.isna(var): return val_str
                sinal = "+" if float(var) >= 0 else "−"
                pct = f"{abs(float(var)) * 100:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".")
                return f"{val_str}<br><span style='font-size:14px; font-weight: bold; color: #ffffff;'>{sinal}{pct} vs Jan-Mai 25</span>"

            base_produtos_acum["Rótulo"] = base_produtos_acum.apply(texto_barra_acum, axis=1)
            threshold_acum = (base_produtos_acum["Valor"].max() if not base_produtos_acum.empty else 1) * 0.20
            text_positions_acum = ["outside" if row["Valor"] < threshold_acum else "inside" for _, row in base_produtos_acum.iterrows()]
            fig_prod_acum = go.Figure(go.Bar(x=base_produtos_acum["Produto"], y=base_produtos_acum["Valor"], text=base_produtos_acum["Rótulo"], textposition=text_positions_acum, textfont=dict(size=16, family="Arial Black", color="#FFFFFF"), insidetextanchor="middle", texttemplate="%{text}", marker_color="#1d6ff2"))
            fig_prod_acum.update_layout(template="plotly_dark", paper_bgcolor="#080f1f", plot_bgcolor="#080f1f", height=390, margin=dict(l=10, r=10, t=10, b=10), showlegend=False, title={"text": "<b>Resultado Contábil acumulado por produto</b>", "font": {"size": 16, "color": "#ffffff"}}, xaxis_showgrid=False, yaxis_showgrid=False, xaxis_zeroline=False, yaxis_zeroline=False)
            fig_prod_acum.update_xaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)")
            fig_prod_acum.update_yaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickprefix="R$ ", separatethousands=True)
            st.plotly_chart(fig_prod_acum, use_container_width=True)

            st.markdown('<div class="section-title">Resumo acumulado das linhas principais por produto</div>', unsafe_allow_html=True)
            st.markdown(tabela_html_pnl_matriz(montar_matriz_pnl_excel(df_pnl_acumulado, obter_linhas_tabela_pnl(df_pnl_acumulado))[0], montar_matriz_pnl_excel(df_pnl_acumulado, obter_linhas_tabela_pnl(df_pnl_acumulado))[1], montar_matriz_pnl_excel(df_pnl_acumulado, obter_linhas_tabela_pnl(df_pnl_acumulado))[2]), unsafe_allow_html=True)
        else:
            st.info(f"Não consegui carregar a aba P&L Acumulado: {erro_pnl_global}")

    with tab_comp_2025:
        try:
            df_comp = carregar_comparativo_2025(arquivo)
            if df_comp is not None and not df_comp.empty:
                df_comp.loc[(df_comp["Ano"] == 2026) & (df_comp["Produto"] == "Total") & (df_comp["Linha_Normalizada"].str.contains("carteira", na=False)), "Realizado"] = 1995139307.0
                
            df_2025_acumulado = carregar_2025_acumulado(arquivo)
            df_comp_principais = montar_comparativo_principais(df_comp, df_2025_acumulado)

            if df_comp.empty or df_comp_principais.empty:
                st.info("Não encontrei dados suficientes na aba Comparativo 2026 x 2025.")
            else:
                st.markdown('<div class="section-title">Comparativo Jan-Mai 26 x Jan-Mai 25</div>', unsafe_allow_html=True)

                novos_cards_linha1 = [
                    ("Margem de Intermediação Jan-Mai 26", "MARGEM INTERMEDIAÇÃO"),
                    ("MG Intermediação Líq. Jan-Mai 26", "MG INTERMEDIAÇÃO LIQ"),
                    ("MG Contribuição Direta Jan-Mai 26", "MG CONTRIBUIÇÃO DIRETA"),
                    ("Resultado Antes do Imposto Jan-Mai 26", "RESULTADO ANTES IMPOSTO"),
                ]
                cols1 = st.columns(4)
                for col, (titulo, linha_nome) in zip(cols1, novos_cards_linha1):
                    with col:
                        linha_df = obter_linha_comparativo(df_comp_principais, linha_nome)
                        if linha_df.empty: card(titulo, 0, ajuda="", variacao=None)
                        else:
                            valor_2025 = linha_df["2025"].iloc[0]
                            valor_2026 = linha_df["2026"].iloc[0]
                            variacao = linha_df["Δ %"].iloc[0]
                            cor_classe = "delta-negative" if variacao is not None and not pd.isna(variacao) and pd.notna(valor_2026) and float(valor_2026) < 0 and float(variacao) < 0 else None
                            variacao_exibir = -variacao if variacao is not None and not pd.isna(variacao) and pd.notna(valor_2026) and float(valor_2026) < 0 and float(variacao) < 0 else variacao
                            card(titulo, valor_2026, ajuda="", variacao=variacao, variacao_label="vs Jan-Mai 25", cor_classe=cor_classe, variacao_exibir=variacao_exibir)

                st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)

                novos_cards_linha2 = [
                    ("Resultado Contábil Jan-Mai 26", "RESULTADO CONTÁBIL"),
                    ("Carteira de Crédito (Média) Jan-Mai 26", "Carteira de Crédito Média"),
                    ("PL Médio Jan-Mai 26", "PL Médio"),
                ]
                cols2 = st.columns(3)
                for col, (titulo, linha_nome) in zip(cols2, novos_cards_linha2):
                    with col:
                        linha_df = obter_linha_comparativo(df_comp_principais, linha_nome)
                        if linha_df.empty: card(titulo, 0, ajuda="", variacao=None)
                        else:
                            valor_2025 = linha_df["2025"].iloc[0]
                            valor_2026 = linha_df["2026"].iloc[0]
                            variacao = linha_df["Δ %"].iloc[0]
                            cor_classe = "delta-negative" if variacao is not None and not pd.isna(variacao) and pd.notna(valor_2026) and float(valor_2026) < 0 and float(variacao) < 0 else None
                            variacao_exibir = -variacao if variacao is not None and not pd.isna(variacao) and pd.notna(valor_2026) and float(valor_2026) < 0 and float(variacao) < 0 else variacao
                            card(titulo, valor_2026, ajuda="", variacao=variacao, variacao_label="vs Jan-Mai 25", cor_classe=cor_classe, variacao_exibir=variacao_exibir)

                st.markdown('<div class="card-row-spacer"></div>', unsafe_allow_html=True)

                st.markdown('<div class="section-title">Quanto do Resultado Contábil acumulado de 2025 já foi alcançado</div>', unsafe_allow_html=True)
                linha_resultado = obter_linha_comparativo(df_comp_principais, "RESULTADO CONTÁBIL")
                if linha_resultado.empty:
                    st.info("Não encontrei a linha de Resultado Contábil para montar o gráfico de alcance.")
                else:
                    valor_base_2025 = linha_resultado["2025 Acumulado"].iloc[0]
                    valor_1t26 = linha_resultado["2026"].iloc[0]
                    fig_alcance_resultado = grafico_alcance_resultado_contabil(valor_1t26, valor_base_2025)
                    if fig_alcance_resultado is not None: st.plotly_chart(fig_alcance_resultado, use_container_width=True)
                    else: st.info("Não foi possível calcular o alcance do Resultado Contábil com a base atual.")

                st.markdown('<div class="section-title">Jan-Mai 25 x Jan-Mai 26 por linha principal</div>', unsafe_allow_html=True)
                termos_ocultos_grafico = ["carteira", "pl medio", "impostos diretos", "comissao", "amortizacao"]
                df_comp_grafico = df_comp_principais[~df_comp_principais["Linha"].map(normalizar_texto).str.contains('|'.join(termos_ocultos_grafico), regex=True, na=False)].copy()
                base_long = df_comp_grafico.melt(id_vars=["Linha", "Ordem"], value_vars=["2025", "2026"], var_name="Ano", value_name="Valor").dropna(subset=["Valor"])
                base_long["Rótulo"] = base_long["Valor"].map(formatar_moeda_curta)
                base_long = base_long.sort_values("Ordem", ascending=False)
                fig_comp_ano = px.bar(base_long, x="Valor", y="Linha", color="Ano", text="Rótulo", orientation="h", barmode="group", labels={"Valor": "", "Linha": "", "Ano": ""})
                fig_comp_ano.update_traces(texttemplate="<b>%{text}</b>", textposition="outside", textfont=dict(size=11, color="#FFFFFF", family="Arial Black"), cliponaxis=False)
                if not base_long.empty:
                    xmin, xmax = base_long["Valor"].min(), base_long["Valor"].max()
                    xpad = max((xmax - xmin) * 0.18, 1)
                    fig_comp_ano.update_xaxes(range=[xmin - xpad, xmax + xpad], showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)", tickprefix="R$ ", separatethousands=True)
                fig_comp_ano.update_yaxes(showgrid=False, zeroline=False, showline=False, gridcolor="rgba(0,0,0,0)")
                fig_comp_ano.update_layout(template="plotly_dark", paper_bgcolor="#080f1f", plot_bgcolor="#080f1f", height=520, margin=dict(l=10, r=120, t=40, b=20), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, font=dict(size=13, color="#ffffff", family="Arial Black")), xaxis_showgrid=False, yaxis_showgrid=False, xaxis_zeroline=False, yaxis_zeroline=False)
                st.plotly_chart(fig_comp_ano, use_container_width=True)

                st.markdown('<div class="section-title">Tabela comparativa</div>', unsafe_allow_html=True)
                termos_ocultos_tabela = ["impostos diretos", "comissao", "amortizacao"]
                df_comp_tabela = df_comp_principais[~df_comp_principais["Linha"].map(normalizar_texto).str.contains('|'.join(termos_ocultos_tabela), regex=True, na=False)].copy()
                st.markdown(tabela_html_comparativo(df_comp_tabela), unsafe_allow_html=True)

        except Exception as erro:
            st.info(f"Não consegui carregar a aba Comparativo 2026 x 2025: {erro}")
else:
    st.info("Aguardando arquivo de base de dados.")
