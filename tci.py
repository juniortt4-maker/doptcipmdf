# -*- coding: utf-8 -*-
"""
Painel Executivo TCI - Termo de Compromisso Institucional
Executar: streamlit run app_painel_tci.py
Requisitos: streamlit pandas openpyxl plotly
"""

import base64
import mimetypes
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.express as px
import streamlit as st


# =====================================
# CONFIGURAÇÕES
# =====================================

ARQUIVO = Path("relatorio-tci.xlsx")

AZUL = "#082C6C"
AZUL_CLARO = "#1D4E89"
VERDE = "#0F766E"
AMBAR = "#B45309"
VERMELHO = "#991B1B"

SEM_INFO = "Não informado"

PLACEHOLDERS = {
    "undefined",
    "null",
    "none",
    "nan",
    "n/a",
    "na",
    "nd",
    "não informado",
    "nao informado",
    "-",
    "--",
    "---",
    "#n/d",
    "?",
    "sem informação",
}

COLUNAS_IGNORADAS = {
    "CPF/CNPJ",
    "LAT",
    "LONG",
    "dataCadastro",
}

st.set_page_config(
    page_title="Painel TCI",
    page_icon="🚔",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =====================================
# CSS
# =====================================

st.markdown(
    """
<style>
.stApp {
    background: #f2f5fa;
}

h1, h2, h3, h4, h5 {
    color: #082C6C !important;
    font-weight: 800 !important;
}

/* Cabeçalho */
.cabecalho-pmdf {
    display: flex;
    align-items: center;
    gap: 24px;
    padding: 22px 28px;
    margin-bottom: 24px;
    border-radius: 20px;
    background: linear-gradient(100deg, #082C6C, #1D4E89);
    box-shadow: 0 6px 18px rgba(8, 44, 108, 0.22);
}

.cabecalho-pmdf .logo-area {
    display: flex;
    align-items: center;
    justify-content: center;
    flex: 0 0 108px;
    width: 108px;
    height: 126px;
    padding: 8px;
    box-sizing: border-box;
    background: #FFFFFF;
    border-radius: 14px;
}

.cabecalho-pmdf .logo-area img {
    display: block;
    width: 100%;
    height: 100%;
    object-fit: contain;
}

.cabecalho-pmdf .cabecalho-texto {
    min-width: 0;
}

.cabecalho-pmdf .instituicao {
    margin: 0 0 6px;
    color: #FFFFFF !important;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
}

.cabecalho-pmdf h1 {
    margin: 0;
    padding: 0;
    color: #FFFFFF !important;
    font-size: clamp(24px, 3vw, 36px);
    font-weight: 800;
    line-height: 1.2;
}

.cabecalho-pmdf .subtitulo {
    margin: 8px 0 0;
    color: #E2E8F0 !important;
    font-size: 16px;
    line-height: 1.4;
}

@media (max-width: 600px) {
    .cabecalho-pmdf {
        flex-direction: column;
        gap: 16px;
        padding: 20px;
        text-align: center;
    }

    .cabecalho-pmdf .logo-area {
        flex: none;
        width: 90px;
        height: 104px;
    }

    .cabecalho-pmdf .instituicao {
        font-size: 10px;
        letter-spacing: 1px;
    }
}

/* Títulos de seção */
.secao {
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 26px 0 14px 0;
    padding: 12px 18px;
    background: #FFFFFF;
    border-left: 8px solid #082C6C;
    border-radius: 14px;
    box-shadow: 0 3px 12px rgba(0, 0, 0, .07);
}

.secao .icone {
    width: 42px;
    height: 42px;
    border-radius: 12px;
    background: #082C6C;
    color: #FFFFFF;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 20px;
    flex: 0 0 42px;
}

.secao .texto {
    font-size: 20px;
    font-weight: 800;
    color: #082C6C;
    letter-spacing: .3px;
    text-transform: none;
}

.secao .texto small {
    display: block;
    font-size: 12px;
    font-weight: 600;
    color: #64748B;
    letter-spacing: .6px;
    text-transform: uppercase;
}

.secao.verde {
    border-left-color: #0F766E;
}

.secao.verde .icone {
    background: #0F766E;
}

.secao.verde .texto {
    color: #0F766E;
}

.secao.ambar {
    border-left-color: #B45309;
}

.secao.ambar .icone {
    background: #B45309;
}

.secao.ambar .texto {
    color: #B45309;
}

.secao.vermelho {
    border-left-color: #991B1B;
}

.secao.vermelho .icone {
    background: #991B1B;
}

.secao.vermelho .texto {
    color: #991B1B;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #082c6c;
}

section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {
    color: #FFFFFF !important;
}

section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea {
    color: #0F172A !important;
}

section[data-testid="stSidebar"] div[data-baseweb="select"] * {
    color: #0F172A !important;
}

section[data-testid="stSidebar"] div[data-baseweb="tag"] span {
    color: #FFFFFF !important;
}

hr {
    margin: 14px 0;
}
</style>
""",
    unsafe_allow_html=True,
)


# =====================================
# CABEÇALHO COM LOGOTIPO PMDF
# =====================================

PASTA_APP = Path(__file__).resolve().parent

EXTENSOES_LOGO = {
    ".jpg",
    ".jpeg",
    ".png",
    ".webp",
}

logotipo = next(
    (
        arquivo
        for arquivo in sorted(PASTA_APP.iterdir())
        if arquivo.is_file()
        and arquivo.stem.strip().casefold() == "brasão so"
        and arquivo.suffix.lower() in EXTENSOES_LOGO
    ),
    None,
)

logo_html = ""

if logotipo is not None:
    imagem_base64 = base64.b64encode(
        logotipo.read_bytes()
    ).decode("ascii")

    tipo_imagem = (
        mimetypes.guess_type(logotipo.name)[0]
        or "image/jpeg"
    )

    logo_html = (
        '<div class="logo-area">'
        f'<img src="data:{tipo_imagem};base64,{imagem_base64}" '
        'alt="Brasão SO - PMDF">'
        '</div>'
    )
else:
    st.warning(
        "Logotipo não encontrado. Verifique se a imagem "
        "'BRASÃO SO' está na mesma pasta do aplicativo, "
        "com extensão .jpg, .jpeg, .png ou .webp."
    )

cabecalho_html = (
    '<div class="cabecalho-pmdf">'
    f'{logo_html}'
    '<div class="cabecalho-texto">'
    '<p class="instituicao">'
    'Polícia Militar do Distrito Federal'
    '</p>'
    '<h1>Painel Executivo TCI</h1>'
    '<p class="subtitulo">'
    'Termo de Costatação de Irregularidades'
    '</p>'
    '</div>'
    '</div>'
)

st.markdown(
    cabecalho_html,
    unsafe_allow_html=True,
)


def titulo_secao(
    texto: str,
    icone: str = "📌",
    subtitulo: str = "",
    cor: str = "",
) -> None:
    """Título de seção com alto contraste."""
    classe = f"secao {cor}".strip()
    sub = f"<small>{subtitulo}</small>" if subtitulo else ""

    st.markdown(
        f"""
        <div class="{classe}">
            <div class="icone">{icone}</div>
            <div class="texto">{texto}{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =====================================
# CARREGAMENTO
# =====================================

@st.cache_data(
    show_spinner="Carregando relatorio-tci.xlsx...",
    ttl=600,
)
def carregar(caminho: str) -> pd.DataFrame:
    df = pd.read_excel(
        caminho,
        dtype=str,
        engine="openpyxl",
    )

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    for col in df.columns:
        if df[col].dtype != object:
            continue

        serie = (
            df[col]
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        eh_lixo = serie.str.lower().isin(PLACEHOLDERS)

        if col in COLUNAS_IGNORADAS:
            df[col] = serie.mask(eh_lixo)
        else:
            df[col] = serie.mask(eh_lixo, SEM_INFO)

    for col in ["LAT", "LONG"]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce",
            )

    if "dataCadastro" in df.columns:
        df["dataCadastro"] = pd.to_datetime(
            df["dataCadastro"],
            errors="coerce",
            dayfirst=True,
        )

    return df


if not ARQUIVO.exists():
    st.error(
        "Arquivo **relatorio-tci.xlsx** não encontrado. "
        "Coloque a planilha na mesma pasta do app "
        "e recarregue a página."
    )
    st.stop()

df = carregar(str(ARQUIVO))

if df.empty:
    st.warning(
        "A planilha foi lida, mas não contém registros."
    )
    st.stop()


# =====================================
# FILTROS
# =====================================

st.sidebar.header("Filtros")

if st.sidebar.button(
    "🧹 Limpar filtros",
    use_container_width=True,
):
    for chave in [
        "f_cidade",
        "f_batalhao",
        "f_status",
        "f_periodo",
        "f_busca",
    ]:
        st.session_state.pop(chave, None)

    st.rerun()

dff = df.copy()

if "Cidade" in df.columns:
    cidades = sorted(
        df["Cidade"].dropna().unique()
    )

    cidade = st.sidebar.multiselect(
        "Cidade",
        cidades,
        key="f_cidade",
        placeholder="Todas",
    )

    if cidade:
        dff = dff[
            dff["Cidade"].isin(cidade)
        ]

if "Batalhão" in dff.columns:
    bpms = sorted(
        dff["Batalhão"].dropna().unique()
    )

    bpm = st.sidebar.multiselect(
        "Batalhão",
        bpms,
        key="f_batalhao",
        placeholder="Todos",
    )

    if bpm:
        dff = dff[
            dff["Batalhão"].isin(bpm)
        ]

if "Status" in dff.columns:
    status_opcoes = sorted(
        dff["Status"].dropna().unique()
    )

    status = st.sidebar.multiselect(
        "Status",
        status_opcoes,
        key="f_status",
        placeholder="Todos",
    )

    if status:
        dff = dff[
            dff["Status"].isin(status)
        ]

if "dataCadastro" in df.columns:
    datas_validas = df["dataCadastro"].dropna()

    if not datas_validas.empty:
        dmin = datas_validas.min().date()
        dmax = datas_validas.max().date()

        periodo = st.sidebar.date_input(
            "Período de cadastro",
            value=(dmin, dmax),
            min_value=dmin,
            max_value=dmax,
            format="DD/MM/YYYY",
            key="f_periodo",
        )

        if (
            isinstance(periodo, (tuple, list))
            and len(periodo) == 2
        ):
            ini, fim = periodo

            dff = dff[
                (dff["dataCadastro"].dt.date >= ini)
                & (dff["dataCadastro"].dt.date <= fim)
            ]

        elif (
            isinstance(periodo, (tuple, list))
            and len(periodo) == 1
        ):
            dff = dff[
                dff["dataCadastro"].dt.date >= periodo[0]
            ]

if "CPF/CNPJ" in dff.columns:
    busca = st.sidebar.text_input(
        "Pesquisar CPF/CNPJ",
        key="f_busca",
    )

    if busca:
        alvo = (
            busca
            .replace(".", "")
            .replace("-", "")
            .replace("/", "")
            .strip()
        )

        dff = dff[
            dff["CPF/CNPJ"]
            .astype(str)
            .str.replace(r"\D", "", regex=True)
            .str.contains(
                alvo,
                na=False,
                case=False,
            )
        ]

if dff.empty:
    st.warning(
        "Nenhum registro encontrado "
        "com os filtros selecionados."
    )
    st.stop()


# =====================================
# HELPERS
# =====================================

def numero(valor) -> str:
    try:
        return f"{int(valor):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "—"


def kpi(
    titulo: str,
    valor,
    icone: str,
    cor: str,
) -> None:
    st.markdown(
        f"""
        <div style="
            background:#FFFFFF;
            border-radius:18px;
            padding:22px;
            box-shadow:0 4px 15px rgba(0,0,0,.08);
            border-left:8px solid {cor};
            min-height:120px;
        ">
            <div style="
                display:flex;
                justify-content:space-between;
                align-items:center;
            ">
                <div>
                    <div style="
                        font-size:13px;
                        font-weight:700;
                        color:#475569;
                        text-transform:uppercase;
                        letter-spacing:1px;
                    ">
                        {titulo}
                    </div>
                    <div style="
                        font-size:36px;
                        font-weight:800;
                        color:#0F172A;
                        margin-top:4px;
                    ">
                        {valor}
                    </div>
                </div>
                <div style="
                    width:65px;
                    height:65px;
                    border-radius:18px;
                    background:{cor};
                    display:flex;
                    align-items:center;
                    justify-content:center;
                    font-size:30px;
                    color:#fff;
                    box-shadow:0 6px 14px rgba(0,0,0,.15);
                ">
                    {icone}
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def mostra(
    figura,
    altura: int = 380,
    legenda: bool = False,
    titulo: str = "",
) -> None:
    figura.update_layout(
        template="plotly_white",
        height=altura,
        margin=dict(
            l=10,
            r=10,
            t=60,
            b=10,
        ),
        showlegend=legenda,
        coloraxis_showscale=False,
        title=dict(
            text=titulo,
            x=0.0,
            xanchor="left",
            font=dict(
                size=16,
                color=AZUL,
                family="Arial",
            ),
        ),
        font=dict(
            color="#0F172A",
            size=13,
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    figura.update_xaxes(
        title=None,
        tickfont=dict(color="#0F172A"),
        linecolor="#CBD5E1",
        gridcolor="#E2E8F0",
    )

    figura.update_yaxes(
        title=None,
        tickfont=dict(color="#0F172A"),
        linecolor="#CBD5E1",
        gridcolor="#E2E8F0",
    )

    st.plotly_chart(
        figura,
        use_container_width=True,
        config={"displayModeBar": False},
    )


def ordena_categoria(
    base: pd.DataFrame,
    coluna: str,
    topo: Optional[int] = None,
) -> pd.DataFrame:
    """Ordena as categorias por quantidade decrescente."""
    base = base.sort_values(
        ["Quantidade", coluna],
        ascending=[False, True],
    )

    if topo:
        base = base.head(topo)

    return base


# =====================================
# KPIs
# =====================================

c1, c2, c3, c4 = st.columns(4)

with c1:
    kpi(
        "Total Registros",
        numero(len(dff)),
        "📋",
        AZUL,
    )

with c2:
    kpi(
        "Cidades",
        (
            numero(dff["Cidade"].nunique())
            if "Cidade" in dff.columns
            else "—"
        ),
        "🏙️",
        VERDE,
    )

with c3:
    kpi(
        "Batalhões",
        (
            numero(dff["Batalhão"].nunique())
            if "Batalhão" in dff.columns
            else "—"
        ),
        "🚔",
        AMBAR,
    )

with c4:
    kpi(
        "Distribuidoras",
        (
            numero(dff["Distribuidora"].nunique())
            if "Distribuidora" in dff.columns
            else "—"
        ),
        "🏢",
        VERMELHO,
    )

st.divider()


# =====================================
# GRÁFICOS: CIDADE E BATALHÃO
# =====================================

g1, g2 = st.columns(2)

with g1:
    titulo_secao(
        "Registros por Cidade",
        "🏙️",
        "Top 15",
        "verde",
    )

    if "Cidade" in dff.columns:
        base = ordena_categoria(
            dff.groupby("Cidade")
            .size()
            .reset_index(name="Quantidade"),
            "Cidade",
            topo=15,
        )

        fig = px.bar(
            base,
            x="Cidade",
            y="Quantidade",
            text="Quantidade",
            color="Quantidade",
            color_continuous_scale=[
                "#A9CBBE",
                VERDE,
            ],
        )

        fig.update_traces(
            textposition="outside",
            cliponaxis=False,
        )

        fig.update_xaxes(
            tickangle=-35,
        )

        mostra(fig)

    else:
        st.info(
            "Coluna Cidade não encontrada na planilha."
        )

with g2:
    titulo_secao(
        "Registros por Batalhão",
        "🚔",
        "Distribuição",
        "ambar",
    )

    if "Batalhão" in dff.columns:
        base = ordena_categoria(
            dff.groupby("Batalhão")
            .size()
            .reset_index(name="Quantidade"),
            "Batalhão",
        )

        fig = px.bar(
            base,
            x="Batalhão",
            y="Quantidade",
            text="Quantidade",
            color="Quantidade",
            color_continuous_scale=[
                "#E7C79B",
                AMBAR,
            ],
        )

        fig.update_traces(
            textposition="outside",
            cliponaxis=False,
        )

        fig.update_xaxes(
            tickangle=-35,
        )

        mostra(fig)

    else:
        st.info(
            "Coluna Batalhão não encontrada na planilha."
        )


# =====================================
# STATUS
# =====================================

titulo_secao(
    "Status dos Registros",
    "📊",
    "Situação atual",
    "",
)

if "Status" in dff.columns:
    status_df = ordena_categoria(
        dff.groupby("Status")
        .size()
        .reset_index(name="Quantidade"),
        "Status",
    )

    fig = px.pie(
        status_df,
        names="Status",
        values="Quantidade",
        hole=0.55,
        color_discrete_sequence=[
            "#082C6C",
            "#1D4E89",
            "#2E6FAD",
            "#5A94CC",
            "#93B8DC",
            "#C0D4E8",
            "#64748B",
        ],
    )

    fig.update_traces(
        textinfo="percent+label",
        textposition="outside",
        textfont=dict(
            size=13,
            color="#0F172A",
        ),
        marker=dict(
            line=dict(
                color="#FFFFFF",
                width=2,
            )
        ),
    )

    fig.update_layout(
        template="plotly_white",
        height=440,
        margin=dict(
            l=10,
            r=10,
            t=30,
            b=10,
        ),
        showlegend=True,
        legend=dict(
            orientation="h",
            y=-0.12,
            font=dict(color="#0F172A"),
        ),
        paper_bgcolor="rgba(0,0,0,0)",
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )

else:
    st.info(
        "Coluna Status não encontrada na planilha."
    )


# =====================================
# TOP DISTRIBUIDORAS
# =====================================

titulo_secao(
    "Top Distribuidoras",
    "🏆",
    "Ranking por volume",
    "vermelho",
)

if "Distribuidora" in dff.columns:
    ranking = ordena_categoria(
        dff.groupby("Distribuidora")
        .size()
        .reset_index(name="Quantidade"),
        "Distribuidora",
        topo=15,
    ).sort_values(
        "Quantidade",
        ascending=True,
    )

    fig = px.bar(
        ranking,
        x="Quantidade",
        y="Distribuidora",
        orientation="h",
        text="Quantidade",
        color="Quantidade",
        color_continuous_scale=[
            "#F3C6C6",
            VERMELHO,
        ],
    )

    fig.update_traces(
        textposition="outside",
        cliponaxis=False,
    )

    fig.update_yaxes(
        automargin=True,
    )

    mostra(
        fig,
        altura=max(400, 30 * len(ranking)),
    )

else:
    st.info(
        "Coluna Distribuidora não encontrada na planilha."
    )


# =====================================
# EVOLUÇÃO TEMPORAL
# =====================================

if "dataCadastro" in dff.columns:
    titulo_secao(
        "Evolução Temporal",
        "📈",
        "Cadastros por dia",
        "verde",
    )

    base = dff.dropna(
        subset=["dataCadastro"]
    )

    if base.empty:
        st.info(
            "Nenhuma data de cadastro válida "
            "nos registros filtrados."
        )

    else:
        evo = (
            base.assign(
                Data=base["dataCadastro"].dt.date
            )
            .groupby("Data")
            .size()
            .reset_index(name="Quantidade")
            .sort_values("Data")
        )

        fig = px.line(
            evo,
            x="Data",
            y="Quantidade",
            markers=True,
            color_discrete_sequence=[AZUL],
        )

        fig.update_traces(
            line=dict(width=3),
            marker=dict(size=7),
        )

        fig.update_yaxes(
            rangemode="tozero",
        )

        mostra(fig)


# =====================================
# DADOS DETALHADOS
# =====================================

titulo_secao(
    "Dados Detalhados",
    "📄",
    f"{numero(len(dff))} registros",
    "",
)

tabela = dff.copy()

if "dataCadastro" in tabela.columns:
    tabela["dataCadastro"] = (
        tabela["dataCadastro"]
        .dt.strftime("%d/%m/%Y")
    )

tabela = tabela.astype(object).where(
    pd.notna(tabela),
    "",
)

st.dataframe(
    tabela,
    use_container_width=True,
    height=700,
    hide_index=True,
)

st.download_button(
    "⬇️ Baixar dados filtrados (CSV)",
    data=(
        dff.to_csv(
            index=False,
            sep=";",
        )
        .encode("utf-8-sig")
    ),
    file_name="tci_filtrado.csv",
    mime="text/csv",
)