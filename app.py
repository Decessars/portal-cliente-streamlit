from __future__ import annotations

import base64
import json
import re
import unicodedata
from datetime import datetime
from html import escape
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
ASSETS_DIR = BASE_DIR / "assets"
ANEXOS_DIR = DATA_DIR / "anexos"
EMPRESAS_DATA_DIR = DATA_DIR / "empresas"
CSV_PATH = DATA_DIR / "dados_demo.csv"
XLSX_PATH = DATA_DIR / "dados_demo.xlsx"
CONFIG_PATH = BASE_DIR / "config_clientes.json"
LOGO_FULL_PATH = ASSETS_DIR / "mh_log_logo_app_512.png"

TEMA_MH = {
    "bg": "#f4faf5",
    "panel": "#ffffff",
    "panel_soft": "#e9f5ec",
    "text": "#173326",
    "muted": "#60776a",
    "accent": "#2f8f5b",
    "accent_alt": "#b8942d",
    "danger": "#dc2626",
    "ok": "#15803d",
    "warning": "#b7791f",
    "border": "#c8ddce",
}

# Tipos monitorados no dominio_dmls_08.py em CONTAS_PASSIVO_PAGAR_CONTROLADAS.
TIPOS_CONTA_PAGAR = {
    "2.1.4.01.003": "ISS A RECOLHER",
    "2.1.4.01.006": "IMPOSTOS LUCRO REAL A RECOLHER",
    "2.1.4.01.015": "SIMPLES NACIONAL A RECOLHER",
    "2.1.5.01.001": "SALARIOS E ORDENADOS A PAGAR",
    "2.1.5.02.001": "INSS A RECOLHER",
    "2.1.5.02.002": "FGTS A RECOLHER",
    "2.1.6.02.001": "HONORARIOS CONTABEIS",
}

COLUNAS_BASE = [
    "empresa",
    "competencia",
    "tipo",
    "descricao",
    "fornecedor_cliente",
    "vencimento",
    "pagamento_recebimento",
    "valor",
    "status",
    "categoria",
    "observacao",
    "documento",
]

COLUNAS_EXTRAS = [
    "tipo_conta_codigo",
    "tipo_conta_nome",
    "anexo_nome",
    "anexo_caminho",
    "codigo_pagamento",
    "criado_em",
    "criado_por",
    "excluido_em",
    "excluido_por",
    "ativo",
]

COLUNAS_DADOS = COLUNAS_BASE + COLUNAS_EXTRAS

COLUNAS_TEXTO = [
    "empresa",
    "competencia",
    "tipo",
    "descricao",
    "fornecedor_cliente",
    "vencimento",
    "pagamento_recebimento",
    "status",
    "categoria",
    "observacao",
    "documento",
    "tipo_conta_codigo",
    "tipo_conta_nome",
    "anexo_nome",
    "anexo_caminho",
    "codigo_pagamento",
    "criado_em",
    "criado_por",
    "excluido_em",
    "excluido_por",
]

COLUNAS_IMPORTACAO = [
    "descricao",
    "fornecedor",
    "vencimento",
    "valor",
    "status",
    "categoria",
    "documento",
    "tipo_conta_codigo",
    "codigo_pagamento",
    "observacao",
]

COLUNAS_IMPORTACAO_OBRIGATORIAS = ["descricao", "fornecedor", "vencimento", "valor"]

DOCUMENTOS_DEMO_LEGADOS = {
    "AP-MHL-001",
    "AP-MHL-002",
    "AP-MHL-003",
    "AP-MHB-001",
    "AP-MHB-002",
}

SUGESTOES_DESCRICAO = [
    "Honorários contábeis",
    "DAS Simples Nacional",
    "INSS a recolher",
    "FGTS a recolher",
    "ISS a recolher",
    "Aluguel",
    "Energia elétrica",
    "Internet",
    "Manutenção de frota",
    "Serviços administrativos",
]

SUGESTOES_FORNECEDOR = [
    "Escritório Contábil",
    "Receita Federal",
    "Prefeitura",
    "Caixa Econômica Federal",
    "Companhia de Energia",
    "Fornecedor Brasil",
]

SUGESTOES_CATEGORIA = [
    "Contabilidade",
    "Tributos",
    "Folha",
    "Administrativo",
    "Operacional",
    "Despesas fixas",
    "Tecnologia",
]

OPCAO_NOVO = "Novo"
OPCOES_OCULTAS = {"outro", "decessars monteiro"}

MODULOS_CONTABEIS = [
    {"id": "contas_a_pagar", "titulo": "Contas a Pagar", "ativo": True},
    {"id": "contas_a_receber", "titulo": "Contas a Receber", "ativo": False},
    {"id": "conciliacao_bancaria", "titulo": "Conciliação Bancária", "ativo": False},
    {"id": "extratos_lancamentos", "titulo": "Extratos e Lançamentos", "ativo": False},
    {"id": "receitas_notas", "titulo": "Receitas e Notas Fiscais", "ativo": False},
    {"id": "impostos_guias", "titulo": "Impostos e Guias", "ativo": False},
    {"id": "folha_inss_fgts", "titulo": "Folha, INSS e FGTS", "ativo": False},
    {"id": "plano_contas", "titulo": "Plano de Contas", "ativo": False},
    {"id": "centros_custo", "titulo": "Centros de Custo", "ativo": False},
    {"id": "balancete", "titulo": "Balancete", "ativo": False},
    {"id": "dre", "titulo": "DRE", "ativo": False},
    {"id": "balanco_patrimonial", "titulo": "Balanço Patrimonial", "ativo": False},
    {"id": "razao_contabil", "titulo": "Razão Contábil", "ativo": False},
    {"id": "relatorios", "titulo": "Relatórios PDF/Excel", "ativo": False},
    {"id": "backup_auditoria", "titulo": "Backup e Auditoria", "ativo": False},
]


def configurar_pagina() -> None:
    st.set_page_config(
        page_title="Portal Contábil do Cliente",
        page_icon="📊",
        layout="wide",
    )

    st.markdown(
        f"""
        <style>
            :root {{
                --mh-bg: {TEMA_MH["bg"]};
                --mh-panel: {TEMA_MH["panel"]};
                --mh-panel-soft: {TEMA_MH["panel_soft"]};
                --mh-text: {TEMA_MH["text"]};
                --mh-muted: {TEMA_MH["muted"]};
                --mh-accent: {TEMA_MH["accent"]};
                --mh-accent-alt: {TEMA_MH["accent_alt"]};
                --mh-border: {TEMA_MH["border"]};
                --mh-danger: {TEMA_MH["danger"]};
                --mh-ok: {TEMA_MH["ok"]};
                --mh-warning: {TEMA_MH["warning"]};
            }}
            [data-testid="stAppViewContainer"] {{
                background:
                    radial-gradient(circle at top left, rgba(47, 143, 91, 0.10), transparent 30rem),
                    linear-gradient(180deg, var(--mh-bg) 0%, #ffffff 72%);
                color: var(--mh-text);
            }}
            [data-testid="stSidebar"] {{
                background: linear-gradient(180deg, #ffffff 0%, var(--mh-panel-soft) 100%);
                border-right: 1px solid var(--mh-border);
            }}
            [data-testid="stSidebar"] img {{
                max-width: 132px;
                display: block;
                margin: 0 auto 0.75rem;
            }}
            .main .block-container {{
                padding-top: 1.25rem;
                padding-bottom: 2rem;
            }}
            .portal-header {{
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 1.25rem;
                padding: 1.1rem 1.25rem;
                background: var(--mh-panel);
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                box-shadow: 0 10px 28px rgba(23, 51, 38, 0.08);
                margin-bottom: 1rem;
            }}
            .portal-brand {{
                display: flex;
                align-items: center;
                gap: 1rem;
                min-width: 0;
            }}
            .portal-logo {{
                width: 150px;
                max-width: 100%;
                height: auto;
                object-fit: contain;
                display: block;
            }}
            .portal-title {{
                margin: 0;
                font-size: 2rem;
                line-height: 1.1;
                color: var(--mh-text);
                font-weight: 800;
                letter-spacing: 0;
            }}
            .portal-subtitle {{
                margin: 0.35rem 0 0;
                color: var(--mh-muted);
                font-size: 0.98rem;
            }}
            .portal-meta {{
                display: flex;
                flex-wrap: wrap;
                justify-content: flex-end;
                gap: 0.55rem;
                min-width: 18rem;
            }}
            .meta-pill {{
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                background: var(--mh-panel-soft);
                color: var(--mh-text);
                padding: 0.55rem 0.7rem;
                font-size: 0.88rem;
                font-weight: 700;
            }}
            .meta-pill span {{
                display: block;
                color: var(--mh-muted);
                font-size: 0.72rem;
                font-weight: 600;
                text-transform: uppercase;
            }}
            .status-ok {{
                border-color: rgba(21, 128, 61, 0.35);
                background: rgba(21, 128, 61, 0.10);
                color: var(--mh-ok);
            }}
            .status-alerta {{
                border-color: rgba(220, 38, 38, 0.35);
                background: rgba(220, 38, 38, 0.10);
                color: var(--mh-danger);
            }}
            .status-pendente {{
                border-color: rgba(183, 121, 31, 0.38);
                background: rgba(183, 121, 31, 0.12);
                color: var(--mh-warning);
            }}
            [data-testid="stMetric"] {{
                background: var(--mh-panel);
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                padding: 16px 18px;
                box-shadow: 0 6px 18px rgba(23, 51, 38, 0.07);
            }}
            [data-testid="stMetricValue"] {{
                color: var(--mh-accent);
            }}
            .metric-grid {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(13rem, 1fr));
                gap: 1rem;
                margin: 1rem 0;
            }}
            .metric-card {{
                background: var(--mh-panel);
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                padding: 1rem 1.15rem;
                min-width: 0;
                box-shadow: 0 6px 18px rgba(23, 51, 38, 0.07);
            }}
            .metric-label {{
                color: var(--mh-muted);
                font-size: 0.9rem;
                font-weight: 700;
                margin-bottom: 0.45rem;
            }}
            .metric-value {{
                color: var(--mh-accent);
                font-size: clamp(1.35rem, 3.2vw, 2rem);
                line-height: 1.15;
                font-weight: 500;
                white-space: nowrap;
            }}
            .metric-delta {{
                display: inline-flex;
                align-items: center;
                width: fit-content;
                margin-top: 0.55rem;
                border-radius: 999px;
                background: rgba(21, 128, 61, 0.10);
                color: var(--mh-ok);
                padding: 0.12rem 0.45rem;
                font-size: 0.85rem;
                font-weight: 800;
            }}
            .login-logo {{
                display: flex;
                justify-content: center;
                margin-bottom: 0.75rem;
            }}
            .login-logo img {{
                width: 150px;
                max-width: 100%;
                height: auto;
                object-fit: contain;
                display: block;
            }}
            .login-panel {{
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                background: var(--mh-panel);
                box-shadow: 0 12px 34px rgba(23, 51, 38, 0.10);
                padding: 1.4rem;
                text-align: center;
            }}
            .login-panel h2 {{
                margin: 0 0 0.35rem;
                font-size: 1.55rem;
            }}
            .login-panel p {{
                color: var(--mh-muted);
                margin: 0.25rem 0;
            }}
            .login-badge {{
                display: inline-flex;
                width: fit-content;
                border-radius: 8px;
                background: rgba(47, 143, 91, 0.12);
                color: var(--mh-accent);
                border: 1px solid rgba(47, 143, 91, 0.24);
                padding: 0.42rem 0.62rem;
                font-size: 0.82rem;
                font-weight: 800;
            }}
            .empresa-card {{
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                background: var(--mh-panel);
                box-shadow: 0 8px 24px rgba(23, 51, 38, 0.06);
                padding: 1rem;
                min-height: 9rem;
            }}
            .empresa-card h3 {{
                margin: 0 0 0.5rem;
                font-size: 1.15rem;
            }}
            .empresa-card p {{
                margin: 0.2rem 0;
                color: var(--mh-muted);
                font-size: 0.9rem;
            }}
            .section-panel {{
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                background: var(--mh-panel);
                box-shadow: 0 8px 24px rgba(23, 51, 38, 0.06);
                padding: 1rem;
                margin: 0.7rem 0 1rem;
            }}
            .small-muted {{
                color: var(--mh-muted);
                font-size: 0.9rem;
            }}
            .menu-contabil-titulo {{
                color: var(--mh-text);
                font-weight: 800;
                font-size: 0.98rem;
                margin: 0.4rem 0 0.25rem;
            }}
            .menu-contabil-ajuda {{
                color: var(--mh-muted);
                font-size: 0.78rem;
                margin: 0 0 0.45rem;
            }}
            div.stButton > button,
            div.stDownloadButton > button,
            div.stFormSubmitButton > button {{
                border-radius: 8px;
                border: 1px solid var(--mh-accent);
                background: var(--mh-accent);
                color: #ffffff;
                font-weight: 700;
                min-height: 2.7rem;
                white-space: normal;
            }}
            div.stButton > button:hover,
            div.stDownloadButton > button:hover,
            div.stFormSubmitButton > button:hover {{
                border-color: var(--mh-accent-alt);
                background: var(--mh-accent-alt);
                color: #ffffff;
            }}
            [data-testid="stDataFrame"] {{
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                overflow: hidden;
            }}
            h1, h2, h3 {{
                color: var(--mh-text);
                letter-spacing: 0;
            }}
            hr {{
                border-color: var(--mh-border);
            }}
            @media (max-width: 760px) {{
                [data-testid="stAppViewContainer"] {{
                    background: linear-gradient(180deg, var(--mh-bg) 0%, #ffffff 78%);
                }}
                .main .block-container {{
                    padding-left: 0.85rem;
                    padding-right: 0.85rem;
                    padding-top: 0.75rem;
                    max-width: 100%;
                }}
                [data-testid="stSidebar"] {{
                    border-right: 0;
                    border-bottom: 1px solid var(--mh-border);
                }}
                [data-testid="stSidebar"] img {{
                    max-width: 104px;
                    margin-bottom: 0.45rem;
                }}
                .portal-header {{
                    align-items: flex-start;
                    flex-direction: column;
                    gap: 0.8rem;
                    padding: 0.95rem;
                }}
                .portal-brand {{
                    align-items: flex-start;
                    flex-direction: column;
                    gap: 0.55rem;
                }}
                .portal-logo {{
                    width: 112px;
                }}
                .portal-meta {{
                    justify-content: flex-start;
                    min-width: 0;
                    width: 100%;
                }}
                .meta-pill {{
                    flex: 1 1 100%;
                    padding: 0.5rem 0.65rem;
                }}
                .portal-title {{
                    font-size: 1.35rem;
                    line-height: 1.2;
                }}
                .portal-subtitle {{
                    font-size: 0.9rem;
                }}
                .login-logo img {{
                    width: 118px;
                }}
                .login-panel {{
                    padding: 1rem;
                }}
                .login-panel h2 {{
                    font-size: 1.25rem;
                }}
                .empresa-card {{
                    min-height: 0;
                    padding: 0.9rem;
                    margin-top: 0.2rem;
                }}
                .empresa-card h3 {{
                    font-size: 1rem;
                }}
                .empresa-card p {{
                    font-size: 0.86rem;
                }}
                [data-testid="stMetric"] {{
                    padding: 0.75rem 0.85rem;
                }}
                [data-testid="stMetricLabel"] {{
                    font-size: 0.78rem;
                }}
                [data-testid="stMetricValue"] {{
                    font-size: 1.15rem;
                }}
                .metric-grid {{
                    grid-template-columns: repeat(auto-fit, minmax(10.5rem, 1fr));
                    gap: 0.75rem;
                }}
                .metric-card {{
                    padding: 0.85rem 0.95rem;
                }}
                .metric-label {{
                    font-size: 0.82rem;
                }}
                .metric-value {{
                    font-size: clamp(1.1rem, 5.2vw, 1.55rem);
                }}
                div.stButton > button,
                div.stDownloadButton > button,
                div.stFormSubmitButton > button {{
                    min-height: 3rem;
                    font-size: 0.95rem;
                }}
                [data-testid="stDataFrame"] {{
                    overflow-x: auto;
                }}
                h1 {{
                    font-size: 1.7rem;
                    line-height: 1.2;
                }}
                h2 {{
                    font-size: 1.35rem;
                }}
                h3 {{
                    font-size: 1.12rem;
                }}
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def imagem_base64(caminho: Path) -> str:
    if not caminho.exists():
        return ""
    return base64.b64encode(caminho.read_bytes()).decode("ascii")


def mostrar_cabecalho(empresa: str, status_geral: str, status_tipo: str) -> None:
    logo_b64 = imagem_base64(LOGO_FULL_PATH)
    logo_html = (
        f'<img class="portal-logo" src="data:image/png;base64,{logo_b64}" alt="MH LOG">'
        if logo_b64
        else '<strong style="color:var(--mh-accent);font-size:1.5rem;">MH LOG</strong>'
    )
    status_classe = {
        "ok": "status-ok",
        "alerta": "status-alerta",
        "pendente": "status-pendente",
    }.get(status_tipo, "status-pendente")

    st.markdown(
        f"""
        <section class="portal-header">
            <div class="portal-brand">
                <div>{logo_html}</div>
                <div>
                    <h1 class="portal-title">Portal Contábil do Cliente</h1>
                    <p class="portal-subtitle">
                        Contas a pagar liberadas para acompanhamento e registro manual.
                    </p>
                </div>
            </div>
            <div class="portal-meta">
                <div class="meta-pill"><span>Cliente</span>{escape(empresa)}</div>
                <div class="meta-pill {status_classe}"><span>Status</span>{escape(status_geral)}</div>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def renderizar_metricas(cards: list[dict[str, object]]) -> None:
    itens = []
    for card in cards:
        delta = card.get("delta")
        delta_html = f'<div class="metric-delta">↑ {escape(str(delta))}</div>' if delta is not None else ""
        itens.append(
            '<div class="metric-card">'
            f'<div class="metric-label">{escape(str(card["label"]))}</div>'
            f'<div class="metric-value">{escape(str(card["value"]))}</div>'
            f"{delta_html}"
            "</div>"
        )

    st.markdown(f'<div class="metric-grid">{"".join(itens)}</div>', unsafe_allow_html=True)


def criar_config_demo() -> dict:
    config = {
        "usuarios": [
            {"usuario": "DMLIMA", "senha": "123456", "perfil": "operador", "empresas": ["MHLOG", "MH BRASIL"]},
            {"usuario": "VICTOR", "senha": "123456", "perfil": "operador", "empresas": ["MHLOG", "MH BRASIL"]},
            {"usuario": "ALEX", "senha": "123456", "perfil": "consulta", "empresas": ["MHLOG", "MH BRASIL"]},
        ],
        "clientes": [
            {"empresa": "MHLOG"},
            {"empresa": "MH BRASIL"},
        ]
    }
    CONFIG_PATH.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
    return config


def normalizar_config(config: dict) -> dict:
    clientes = config.get("clientes", [])
    empresas = [cliente.get("empresa", "") for cliente in clientes if cliente.get("empresa")]
    usuarios = config.get("usuarios", [])

    if not usuarios:
        usuarios = [
            {
                "usuario": item.get("empresa", ""),
                "senha": item.get("senha", ""),
                "perfil": "operador",
                "empresas": empresas,
            }
            for item in clientes
            if item.get("senha")
        ]

    for usuario in usuarios:
        usuario["perfil"] = str(usuario.get("perfil", "operador") or "operador").strip().lower()

    for cliente in clientes:
        cliente.pop("senha", None)

    return {"usuarios": usuarios, "clientes": clientes}


def carregar_config() -> dict:
    if not CONFIG_PATH.exists():
        return criar_config_demo()

    with CONFIG_PATH.open("r", encoding="utf-8") as arquivo:
        return normalizar_config(json.load(arquivo))


def dados_demo() -> pd.DataFrame:
    return pd.DataFrame(columns=COLUNAS_DADOS)


def criar_dados_demo() -> pd.DataFrame:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    df = dados_demo()
    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    return df


def normalizar_ativo(valor: object) -> bool:
    if isinstance(valor, bool):
        return valor
    texto = str(valor).strip().lower()
    return texto not in {"false", "0", "nao", "não", "n", "excluido", "excluído"}


def normalizar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.loc[:, ~df.columns.duplicated()].copy()
    for coluna in COLUNAS_DADOS:
        if coluna not in df.columns:
            df[coluna] = True if coluna == "ativo" else ""

    df = df[COLUNAS_DADOS].copy()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0)
    for coluna in COLUNAS_TEXTO:
        df[coluna] = df[coluna].fillna("").astype(str)
    df["ativo"] = df["ativo"].apply(normalizar_ativo)
    df["vencimento_dt"] = pd.to_datetime(df["vencimento"], errors="coerce")
    df["pagamento_recebimento_dt"] = pd.to_datetime(df["pagamento_recebimento"], errors="coerce")
    return df


def slug_empresa(empresa: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "_", empresa.strip().upper()).strip("_")
    return slug or "EMPRESA"


def pasta_empresa(empresa: str) -> Path:
    return EMPRESAS_DATA_DIR / slug_empresa(empresa)


def caminho_dados_empresa(empresa: str) -> Path:
    return pasta_empresa(empresa) / "dados_portal.csv"


def carregar_base_demo() -> pd.DataFrame:
    if CSV_PATH.exists():
        return pd.read_csv(CSV_PATH)
    if XLSX_PATH.exists():
        return pd.read_excel(XLSX_PATH)
    return criar_dados_demo()


def carregar_dados_empresa(empresa: str) -> pd.DataFrame:
    """Carrega a base operacional isolada de uma empresa."""
    caminho = caminho_dados_empresa(empresa)
    if caminho.exists():
        df = pd.read_csv(caminho)
    else:
        demo = carregar_base_demo()
        df = demo.loc[demo["empresa"].astype(str) == empresa].copy()
        salvar_dados_empresa(normalizar_dataframe(df), empresa)
    df_normalizado = normalizar_dataframe(df)
    df_limpo = remover_lancamentos_demo_legados(df_normalizado)
    if len(df_limpo) != len(df_normalizado):
        salvar_dados_empresa(df_limpo, empresa)
    return df_limpo


def remover_lancamentos_demo_legados(df: pd.DataFrame) -> pd.DataFrame:
    documentos = df["documento"].astype(str).str.strip()
    criado_em = df["criado_em"].fillna("").astype(str).str.strip()
    mascara_demo = documentos.isin(DOCUMENTOS_DEMO_LEGADOS) & criado_em.eq("")
    if not mascara_demo.any():
        return df
    return df.loc[~mascara_demo].copy()


def carregar_dados_empresas(config: dict) -> pd.DataFrame:
    frames = [carregar_dados_empresa(cliente["empresa"]) for cliente in config["clientes"]]
    if not frames:
        return normalizar_dataframe(pd.DataFrame(columns=COLUNAS_DADOS))
    return normalizar_dataframe(pd.concat(frames, ignore_index=True))


def salvar_dados_empresa(df: pd.DataFrame, empresa: str) -> Path:
    """Salva no arquivo operacional da empresa selecionada."""
    caminho = caminho_dados_empresa(empresa)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    dados = normalizar_dataframe(df)
    dados = dados.loc[dados["empresa"].astype(str).eq(empresa)].copy()
    dados = dados[[coluna for coluna in COLUNAS_DADOS if coluna in dados.columns]].copy()
    dados.to_csv(caminho, index=False, encoding="utf-8-sig")
    return caminho


def formatar_moeda_br(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_data_br(valor: object) -> str:
    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return ""
    return data.strftime("%d/%m/%Y")


def agora_br() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def gerar_csv_download(df: pd.DataFrame) -> bytes:
    colunas = [coluna for coluna in COLUNAS_DADOS if coluna in df.columns]
    return df[colunas].to_csv(index=False).encode("utf-8-sig")


def gerar_excel_contas_download(df: pd.DataFrame, empresa: str) -> bytes:
    dados = df.sort_values(["vencimento_dt", "descricao"], ascending=[False, True], na_position="last").copy()
    hoje = pd.Timestamp(datetime.now().date())

    colunas = [
        ("vencimento", "Vencimento"),
        ("descricao", "Descricao"),
        ("fornecedor_cliente", "Fornecedor"),
        ("tipo_conta_nome", "Tipo de conta"),
        ("valor", "Valor"),
        ("status", "Status"),
        ("categoria", "Categoria"),
        ("documento", "Documento"),
        ("codigo_pagamento", "Codigo de barras / Pix"),
        ("observacao", "Observacao"),
        ("criado_em", "Incluido em"),
        ("criado_por", "Incluido por"),
    ]
    colunas = [(campo, titulo) for campo, titulo in colunas if campo in dados.columns]

    wb = Workbook()
    ws = wb.active
    ws.title = "Contas a pagar"
    ws.sheet_view.showGridLines = False

    cor_titulo = "173326"
    cor_verde = "2F8F5B"
    cor_verde_claro = "E9F5EC"
    cor_borda = "C8DDCE"
    cor_alerta = "FFF7E6"
    cor_vencida = "FEE2E2"
    cor_texto = "173326"

    ws.merge_cells("A1:L1")
    ws["A1"] = "Contas a pagar"
    ws["A1"].font = Font(bold=True, size=20, color=cor_titulo)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:L2")
    ws["A2"] = f"Empresa: {empresa} | Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(size=10, color="60776A")

    total = float(pd.to_numeric(dados.get("valor", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    vencidas = int(dados.apply(conta_vencida, axis=1).sum()) if not dados.empty else 0
    vence_hoje = int((dados.get("vencimento_dt", pd.Series(dtype="datetime64[ns]")).dt.date == hoje.date()).sum()) if "vencimento_dt" in dados else 0
    proximos_7 = 0
    if "vencimento_dt" in dados:
        proximos_7 = int(((dados["vencimento_dt"] > hoje) & (dados["vencimento_dt"] <= hoje + pd.Timedelta(days=7))).sum())

    resumo = [
        ("Total a pagar", total, "moeda"),
        ("Contas abertas", len(dados), "numero"),
        ("Vencidas", vencidas, "numero"),
        ("Vence hoje", vence_hoje, "numero"),
        ("Proximos 7 dias", proximos_7, "numero"),
    ]
    for indice, (rotulo, valor, tipo) in enumerate(resumo):
        coluna = 1 + (indice * 2)
        ws.merge_cells(start_row=4, start_column=coluna, end_row=4, end_column=coluna + 1)
        ws.merge_cells(start_row=5, start_column=coluna, end_row=5, end_column=coluna + 1)
        cel_rotulo = ws.cell(row=4, column=coluna, value=rotulo)
        cel_valor = ws.cell(row=5, column=coluna, value=valor)
        for linha in (4, 5):
            cel = ws.cell(row=linha, column=coluna)
            cel.fill = PatternFill("solid", fgColor=cor_verde_claro)
            cel.border = Border(
                left=Side(style="thin", color=cor_borda),
                right=Side(style="thin", color=cor_borda),
                top=Side(style="thin", color=cor_borda),
                bottom=Side(style="thin", color=cor_borda),
            )
            cel.alignment = Alignment(horizontal="center", vertical="center")
        cel_rotulo.font = Font(bold=True, size=9, color="60776A")
        cel_valor.font = Font(bold=True, size=14, color=cor_titulo)
        if tipo == "moeda":
            cel_valor.number_format = '"R$" #,##0.00'

    linha_cabecalho = 8
    for col_idx, (_, titulo) in enumerate(colunas, start=1):
        cel = ws.cell(row=linha_cabecalho, column=col_idx, value=titulo)
        cel.fill = PatternFill("solid", fgColor=cor_verde)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = Border(bottom=Side(style="thin", color=cor_borda))

    for row_idx, (_, linha) in enumerate(dados.iterrows(), start=linha_cabecalho + 1):
        nivel = nivel_prazo_conta(linha)
        fill = None
        if nivel == "vencida":
            fill = PatternFill("solid", fgColor=cor_vencida)
        elif nivel == "atencao":
            fill = PatternFill("solid", fgColor=cor_alerta)

        for col_idx, (campo, _) in enumerate(colunas, start=1):
            valor = linha.get(campo, "")
            cel = ws.cell(row=row_idx, column=col_idx)
            if campo in {"vencimento", "criado_em"}:
                data = pd.to_datetime(valor, errors="coerce")
                cel.value = data.to_pydatetime() if not pd.isna(data) else ""
                cel.number_format = "dd/mm/yyyy"
            elif campo == "valor":
                cel.value = float(pd.to_numeric(valor, errors="coerce") or 0)
                cel.number_format = '"R$" #,##0.00'
            else:
                cel.value = str(valor or "")
            cel.font = Font(color=cor_texto)
            cel.alignment = Alignment(vertical="top", wrap_text=True)
            cel.border = Border(bottom=Side(style="hair", color=cor_borda))
            if fill:
                cel.fill = fill

    ultima_linha = max(linha_cabecalho + 1, linha_cabecalho + len(dados))
    ultima_coluna = max(1, len(colunas))
    tabela_ref = f"A{linha_cabecalho}:{get_column_letter(ultima_coluna)}{ultima_linha}"
    tabela = Table(displayName="TabelaContasPagar", ref=tabela_ref)
    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium4",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tabela)

    larguras = {
        "Vencimento": 13,
        "Descricao": 32,
        "Fornecedor": 28,
        "Tipo de conta": 30,
        "Valor": 15,
        "Status": 14,
        "Categoria": 20,
        "Documento": 18,
        "Codigo de barras / Pix": 34,
        "Observacao": 34,
        "Incluido em": 16,
        "Incluido por": 14,
    }
    for col_idx, (_, titulo) in enumerate(colunas, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = larguras.get(titulo, 18)
    ws.freeze_panes = "A9"
    ws.auto_filter.ref = tabela_ref

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    saida = BytesIO()
    wb.save(saida)
    return saida.getvalue()


def opcoes_com_historico(df: pd.DataFrame, coluna: str, padroes: list[str]) -> list[str]:
    opcoes = [""]
    if coluna in df.columns:
        historico = (
            df[coluna]
            .dropna()
            .astype(str)
            .map(str.strip)
            .loc[lambda serie: serie.ne("")]
            .drop_duplicates()
            .tolist()
        )
        opcoes.extend(historico)
    opcoes.extend(padroes)
    unicas = []
    chaves_vistas = set()
    for opcao in opcoes:
        texto = str(opcao or "").strip()
        chave = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").casefold()
        chave = re.sub(r"\s+", " ", chave).strip()
        if chave in OPCOES_OCULTAS:
            continue
        if chave in chaves_vistas:
            continue
        chaves_vistas.add(chave)
        unicas.append(texto)
    return unicas


def resolver_opcao_digitavel(selecao: str, texto_novo: str) -> str:
    if selecao == OPCAO_NOVO:
        return texto_novo.strip()
    return str(selecao or "").strip()


def opcoes_com_novo(opcoes: list[str]) -> list[str]:
    limpas = [
        str(opcao or "").strip()
        for opcao in opcoes
        if str(opcao or "").strip() and str(opcao or "").strip() != OPCAO_NOVO
    ]
    ordenadas = sorted(
        limpas,
        key=lambda texto: unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").casefold(),
    )
    return [OPCAO_NOVO, *ordenadas]


def rotulo_opcao_nova(opcao: str) -> str:
    return f"🔴 {OPCAO_NOVO}" if opcao == OPCAO_NOVO else opcao


def indice_padrao_sem_novo(opcoes: list[str]) -> int:
    for indice, opcao in enumerate(opcoes):
        if str(opcao or "").strip() and opcao != OPCAO_NOVO:
            return indice
    return 0


def campo_opcao_nova(
    container: object,
    label: str,
    opcoes: list[str],
    label_novo: str,
    key_prefix: str,
    index: int | None = None,
) -> tuple[str, str]:
    indice = indice_padrao_sem_novo(opcoes) if index is None else index
    selecao = container.selectbox(
        label,
        opcoes,
        index=indice,
        key=f"{key_prefix}_selecao",
        format_func=rotulo_opcao_nova,
    )
    texto_novo = ""
    if selecao == OPCAO_NOVO:
        texto_novo = container.text_input(label_novo, key=f"{key_prefix}_novo")
    return selecao, texto_novo


def gerar_modelo_importacao_csv() -> bytes:
    exemplo = pd.DataFrame(
        [
            {
                "descricao": "Boleto fornecedor exemplo",
                "fornecedor": "Fornecedor Exemplo",
                "vencimento": "2026-06-10",
                "valor": "1500,00",
                "status": "aberto",
                "categoria": "Administrativo",
                "documento": "NF-0001",
                "tipo_conta_codigo": "2.1.6.02.001",
                "codigo_pagamento": "34191.79001 01043.510047 91020.150008 8 123400000150000",
                "observacao": "Linha de exemplo; apague antes de importar dados reais.",
            }
        ],
        columns=COLUNAS_IMPORTACAO,
    )
    return exemplo.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")


def ler_csv_importacao(arquivo: object) -> pd.DataFrame:
    try:
        return pd.read_csv(arquivo, sep=None, engine="python", dtype=str).fillna("")
    except Exception as erro:
        raise ValueError(f"Não foi possível ler o CSV: {erro}") from erro


def normalizar_valor_importado(valor: object) -> float:
    texto = str(valor).strip()
    if not texto:
        return 0.0
    texto = re.sub(r"[^0-9,.-]", "", texto)
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def parse_valor_br(valor: object) -> float | None:
    texto = str(valor or "").strip()
    if not texto:
        return None

    texto = re.sub(r"[^0-9,.-]", "", texto)
    if not texto:
        return None
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "." in texto:
        texto = texto.replace(".", "")

    try:
        return float(texto)
    except ValueError:
        return None


def preparar_importacao_contas(df_importado: pd.DataFrame, empresa: str, usuario: str) -> tuple[pd.DataFrame, list[str]]:
    df_importado.columns = [str(coluna).strip().lower() for coluna in df_importado.columns]
    faltantes = [coluna for coluna in COLUNAS_IMPORTACAO_OBRIGATORIAS if coluna not in df_importado.columns]
    if faltantes:
        raise ValueError(f"Colunas obrigatórias ausentes: {', '.join(faltantes)}.")

    linhas = []
    erros = []
    for numero_linha, linha in df_importado.iterrows():
        linha_numero = numero_linha + 2
        descricao = str(linha.get("descricao", "")).strip()
        fornecedor = str(linha.get("fornecedor", "")).strip()
        vencimento = pd.to_datetime(linha.get("vencimento", ""), errors="coerce", dayfirst=True)
        valor = normalizar_valor_importado(linha.get("valor", ""))
        status = str(linha.get("status", "aberto")).strip().lower() or "aberto"
        status = status if status in {"aberto", "pendente", "vencido"} else "aberto"
        tipo_codigo = str(linha.get("tipo_conta_codigo", "")).strip()
        tipo_nome = TIPOS_CONTA_PAGAR.get(tipo_codigo, "")
        codigo_pagamento = str(linha.get("codigo_pagamento", "")).strip()

        if not descricao or not fornecedor or pd.isna(vencimento) or valor <= 0:
            erros.append(f"Linha {linha_numero}: preencha descricao, fornecedor, vencimento válido e valor maior que zero.")
            continue

        documento = str(linha.get("documento", "")).strip()
        if not documento:
            documento = f"IMP-{empresa}-{datetime.now().strftime('%Y%m%d%H%M%S')}-{linha_numero}"

        linhas.append(
            {
                "empresa": empresa,
                "competencia": "",
                "tipo": "conta_a_pagar",
                "descricao": descricao,
                "fornecedor_cliente": fornecedor,
                "vencimento": vencimento.strftime("%Y-%m-%d"),
                "pagamento_recebimento": "",
                "valor": valor,
                "status": status,
                "categoria": str(linha.get("categoria", "")).strip(),
                "observacao": str(linha.get("observacao", "")).strip(),
                "documento": documento,
                "tipo_conta_codigo": tipo_codigo,
                "tipo_conta_nome": tipo_nome,
                "anexo_nome": "",
                "anexo_caminho": "",
                "codigo_pagamento": codigo_pagamento,
                "criado_em": agora_br(),
                "criado_por": usuario,
                "excluido_em": "",
                "excluido_por": "",
                "ativo": True,
            }
        )

    return pd.DataFrame(linhas, columns=COLUNAS_DADOS), erros


def validar_login(config: dict, usuario: str, senha_digitada: str) -> bool:
    senha_configurada = obter_senha_usuario(config, usuario)
    return bool(senha_configurada and senha_digitada and senha_digitada == senha_configurada)


def obter_cliente(config: dict, empresa: str) -> dict:
    return next((item for item in config["clientes"] if item["empresa"] == empresa), {})


def obter_usuario(config: dict, usuario: str) -> dict:
    usuario_normalizado = usuario.strip().upper()
    return next(
        (item for item in config["usuarios"] if item.get("usuario", "").strip().upper() == usuario_normalizado),
        {},
    )


def obter_senha_usuario(config: dict, usuario: str) -> str:
    usuario_config = obter_usuario(config, usuario)
    usuario_chave = usuario_config.get("usuario", usuario).strip().upper()
    try:
        usuarios_secret = st.secrets.get("usuarios", {})
        if usuario_chave in usuarios_secret:
            return str(usuarios_secret[usuario_chave])
    except Exception:
        pass

    return str(usuario_config.get("senha", ""))


def perfil_do_usuario(config: dict, usuario: str) -> str:
    return str(obter_usuario(config, usuario).get("perfil", "operador") or "operador").strip().lower()


def usuario_pode_alterar(config: dict, usuario: str) -> bool:
    return perfil_do_usuario(config, usuario) not in {"consulta", "consultor", "visualizacao", "visualização", "somente_consulta"}


def empresas_do_usuario(config: dict, usuario: str) -> list[str]:
    usuario_config = obter_usuario(config, usuario)
    empresas_liberadas = usuario_config.get("empresas", [])
    empresas_configuradas = [cliente["empresa"] for cliente in config["clientes"]]
    if not empresas_liberadas:
        return empresas_configuradas
    return [empresa for empresa in empresas_configuradas if empresa in empresas_liberadas]


def inicializar_sessao() -> None:
    st.session_state.setdefault("autenticado", False)
    st.session_state.setdefault("usuario_logado", "")
    st.session_state.setdefault("empresa_logada", "")
    st.session_state.setdefault("modulo_atual", "contas_a_pagar")


def tela_login(config: dict) -> None:
    logo_b64 = imagem_base64(LOGO_FULL_PATH)
    logo_html = (
        f'<img src="data:image/png;base64,{logo_b64}" alt="MH LOG">'
        if logo_b64
        else '<strong style="color:var(--mh-accent);font-size:1.5rem;">MH LOG</strong>'
    )

    st.write("")
    _, col_login, _ = st.columns([1, 1.05, 1])
    with col_login:
        st.markdown(
            f"""
            <div class="login-logo">{logo_html}</div>
            <section class="login-panel">
                <span class="login-badge">Acesso restrito</span>
                <h2>Portal Contábil do Cliente</h2>
                <p>Contas a pagar liberadas pelo escritório.</p>
                <p>Use seu usuário e senha de acesso.</p>
            </section>
            """,
            unsafe_allow_html=True,
        )
        st.write("")
        with st.container(border=True):
            st.subheader("Entrar no portal")
            st.caption("Informe seus dados para continuar.")
            with st.form("form_login", clear_on_submit=False):
                empresa = st.selectbox("Usuário", empresas, index=empresa_idx)
                senha = st.text_input("Senha de acesso", type="password")
                entrar = st.form_submit_button("Entrar", use_container_width=True)

            if entrar:
                cliente_atual = obter_cliente(config, empresa)
                competencias = cliente_atual.get("competencias", [])
                if validar_login(config, empresa, senha):
                    st.session_state.autenticado = True
                    st.session_state.empresa_logada = empresa
                    st.session_state.competencia_logada = competencias[0] if competencias else ""
                    st.session_state.empresa_login = empresa
                    st.rerun()
                else:
                    st.error("Usuário ou senha inválida.")


def logout() -> None:
    st.session_state.autenticado = False
    st.session_state.empresa_logada = ""
    st.session_state.competencia_logada = ""
    st.rerun()


def sidebar_filtros_antigo(config: dict) -> tuple[str, str]:
    if LOGO_FULL_PATH.exists():
        st.sidebar.image(str(LOGO_FULL_PATH), width=210)
    st.sidebar.title("🔐 Sessão do Cliente")

    empresa = st.session_state.get("empresa_logada", "")
    st.sidebar.text_input("Usuário", value=empresa, disabled=True)
    st.sidebar.divider()
    st.sidebar.success("Acesso liberado")
    st.sidebar.button("Sair", use_container_width=True, on_click=logout)
    return empresa, ""


def filtrar_dados(df: pd.DataFrame, empresa: str) -> pd.DataFrame:
    filtro = df["empresa"] == empresa
    return df.loc[filtro].copy()


def filtrar_contas_a_pagar_abertas(df: pd.DataFrame, empresa: str) -> pd.DataFrame:
    dados = filtrar_dados(df, empresa)
    status_abertos = dados["status"].astype(str).str.lower().isin(["aberto", "pendente", "vencido"])
    filtro = (dados["tipo"] == "conta_a_pagar") & status_abertos & dados["ativo"]
    return dados.loc[filtro].copy()


def tela_login(config: dict) -> None:
    logo_b64 = imagem_base64(LOGO_FULL_PATH)
    logo_html = (
        f'<img src="data:image/png;base64,{logo_b64}" alt="MH LOG">'
        if logo_b64
        else '<strong style="color:var(--mh-accent);font-size:1.5rem;">MH LOG</strong>'
    )

    st.write("")
    _, col_login, _ = st.columns([1, 1.05, 1])
    with col_login:
        st.markdown(
            f"""
            <div class="login-logo">{logo_html}</div>
            <section class="login-panel">
                <span class="login-badge">Acesso restrito</span>
                <h2>Portal Contabil do Cliente</h2>
                <p>Contas a pagar liberadas pelo escritorio.</p>
                <p>Use seu usuario e senha de acesso.</p>
            </section>
            """,
            unsafe_allow_html=True,
        )
        st.write("")
        with st.container(border=True):
            st.subheader("Entrar no portal")
            st.caption("Informe seus dados para continuar.")
            with st.form("form_login", clear_on_submit=False):
                usuario = st.text_input("Usuario", value=st.session_state.get("usuario_login", ""))
                senha = st.text_input("Senha de acesso", type="password")
                entrar = st.form_submit_button("Entrar", use_container_width=True)

            if entrar:
                usuario = usuario.strip().upper()
                empresas_liberadas = empresas_do_usuario(config, usuario)
                if validar_login(config, usuario, senha) and empresas_liberadas:
                    st.session_state.autenticado = True
                    st.session_state.usuario_logado = usuario
                    st.session_state.usuario_login = usuario
                    st.session_state.empresa_logada = ""
                    st.rerun()
                else:
                    st.error("Usuario ou senha invalida.")


def logout() -> None:
    st.session_state.autenticado = False
    st.session_state.usuario_logado = ""
    st.session_state.empresa_logada = ""


def voltar_dashboard_empresas() -> None:
    st.session_state.empresa_logada = ""


def selecionar_empresa(empresa: str, config: dict) -> None:
    st.session_state.empresa_logada = empresa
    st.session_state.modulo_atual = "contas_a_pagar"


def selecionar_modulo(modulo_id: str) -> None:
    st.session_state.modulo_atual = modulo_id


def exibir_menu_contabil() -> None:
    st.sidebar.markdown(
        """
        <div class="menu-contabil-titulo">Menu contábil</div>
        <div class="menu-contabil-ajuda">Os módulos desativados serão liberados aos poucos.</div>
        """,
        unsafe_allow_html=True,
    )
    modulo_atual = st.session_state.get("modulo_atual", "contas_a_pagar")
    for modulo in MODULOS_CONTABEIS:
        ativo = bool(modulo["ativo"])
        selecionado = modulo["id"] == modulo_atual
        label = modulo["titulo"] if ativo else f"{modulo['titulo']} - em breve"
        st.sidebar.button(
            label,
            key=f"menu_{modulo['id']}",
            use_container_width=True,
            disabled=not ativo,
            type="primary" if selecionado and ativo else "secondary",
            on_click=selecionar_modulo if ativo else None,
            args=(modulo["id"],) if ativo else None,
        )


def sidebar_filtros(config: dict) -> tuple[str, str]:
    if LOGO_FULL_PATH.exists():
        st.sidebar.image(str(LOGO_FULL_PATH), width=120)

    usuario = st.session_state.get("usuario_logado", "")
    empresa = st.session_state.get("empresa_logada", "")

    exibir_menu_contabil()
    st.sidebar.divider()
    st.sidebar.title("Sessao")
    st.sidebar.text_input("Usuario", value=usuario, disabled=True)
    st.sidebar.text_input("Empresa", value=empresa, disabled=True)
    st.sidebar.text_input("Perfil", value=perfil_do_usuario(config, usuario).title(), disabled=True)
    st.sidebar.divider()
    st.sidebar.success("Acesso liberado")
    st.sidebar.button("Trocar empresa", use_container_width=True, on_click=voltar_dashboard_empresas)
    st.sidebar.button("Sair", use_container_width=True, on_click=logout)
    return empresa, usuario


def dashboard_empresas(config: dict, df: pd.DataFrame) -> None:
    usuario = st.session_state.get("usuario_logado", "")
    empresas = empresas_do_usuario(config, usuario)

    if LOGO_FULL_PATH.exists():
        st.sidebar.image(str(LOGO_FULL_PATH), width=120)
    st.sidebar.markdown(
        """
        <div class="menu-contabil-titulo">Menu contábil</div>
        <div class="menu-contabil-ajuda">Escolha uma empresa para liberar os módulos.</div>
        """,
        unsafe_allow_html=True,
    )
    for modulo in MODULOS_CONTABEIS:
        label = modulo["titulo"] if modulo["ativo"] else f"{modulo['titulo']} - em breve"
        st.sidebar.button(label, key=f"dashboard_menu_{modulo['id']}", use_container_width=True, disabled=True)
    st.sidebar.divider()
    st.sidebar.title("Sessao")
    st.sidebar.text_input("Usuario", value=usuario, disabled=True)
    st.sidebar.button("Sair", use_container_width=True, on_click=logout)

    st.title("Dashboard por empresa")
    st.caption("Escolha uma empresa para abrir as contas a pagar e os indicadores dela.")

    cols = st.columns(2, gap="medium")
    for indice, empresa in enumerate(empresas):
        contas_empresa = df.loc[
            (df["empresa"] == empresa)
            & (df["tipo"] == "conta_a_pagar")
            & df["status"].astype(str).str.lower().isin(["aberto", "pendente", "vencido"])
            & df["ativo"]
        ]
        total = contas_empresa["valor"].sum()
        vencidas = int(contas_empresa.apply(conta_vencida, axis=1).sum())

        with cols[indice % 2]:
            st.markdown(
                f"""
                <div class="empresa-card">
                    <h3>{escape(empresa)}</h3>
                    <p>Total em aberto: <strong>{escape(formatar_moeda_br(total))}</strong></p>
                    <p>Contas abertas: <strong>{len(contas_empresa)}</strong></p>
                    <p>Vencidas: <strong>{vencidas}</strong></p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.button(
                f"Abrir {empresa}",
                key=f"abrir_empresa_{empresa}",
                use_container_width=True,
                on_click=selecionar_empresa,
                args=(empresa, config),
            )


def status_geral_contas(df: pd.DataFrame) -> tuple[str, str]:
    if df.empty:
        return "Sem contas em aberto", "ok"

    vencidos = int(df.apply(conta_vencida, axis=1).sum())
    if vencidos:
        return f"{vencidos} conta(s) vencida(s)", "alerta"
    return f"{len(df)} conta(s) em aberto", "pendente"


def conta_vencida(linha: pd.Series) -> bool:
    status = str(linha.get("status", linha.get("Status", ""))).strip().lower()
    if status == "vencido":
        return True
    if status not in {"aberto", "pendente"}:
        return False

    vencimento_valor = linha.get("vencimento", linha.get("Vencimento", ""))
    vencimento_texto = str(vencimento_valor)
    vencimento = pd.to_datetime(vencimento_valor, errors="coerce", dayfirst="/" in vencimento_texto)
    if pd.isna(vencimento):
        return False
    return vencimento.date() < datetime.now().date()


def nivel_prazo_conta(linha: pd.Series) -> str:
    if conta_vencida(linha):
        return "vencida"

    status = str(linha.get("status", linha.get("Status", ""))).strip().lower()
    if status not in {"aberto", "pendente"}:
        return "neutra"

    vencimento_valor = linha.get("vencimento", linha.get("Vencimento", ""))
    vencimento_texto = str(vencimento_valor)
    vencimento = pd.to_datetime(vencimento_valor, errors="coerce", dayfirst="/" in vencimento_texto)
    if pd.isna(vencimento):
        return "neutra"

    dias = (vencimento.date() - datetime.now().date()).days
    if dias <= 7:
        return "atencao"
    return "ok"


def indicador_prazo_conta(linha: pd.Series) -> str:
    nivel = nivel_prazo_conta(linha)
    if nivel == "vencida":
        return "🔴"
    if nivel == "atencao":
        return "🟡"
    if nivel == "ok":
        return "🟢"
    return "⚪"


def escrever_celula_conta(coluna: object, texto: str, nivel: str, nowrap: bool = False, indicador: str = "") -> None:
    cor = {"vencida": "#b91c1c", "atencao": "#b7791f", "ok": "inherit"}.get(nivel, "inherit")
    peso = "700" if nivel == "vencida" else "400"
    quebra = "white-space:nowrap;" if nowrap else "overflow-wrap:anywhere;"
    conteudo = f"{indicador} {texto}".strip()
    coluna.markdown(
        f'<span style="color:{cor};font-weight:{peso};{quebra}">{escape(conteudo)}</span>',
        unsafe_allow_html=True,
    )


def preparar_tabela_contas(df: pd.DataFrame) -> pd.DataFrame:
    colunas = [
        "vencimento",
        "descricao",
        "fornecedor_cliente",
        "tipo_conta_nome",
        "valor",
        "status",
        "categoria",
        "documento",
        "anexo_nome",
        "codigo_pagamento",
        "criado_em",
        "criado_por",
    ]
    dados = df.sort_values(["vencimento_dt", "descricao"], ascending=[False, True], na_position="last").copy()
    dados = dados[[coluna for coluna in colunas if coluna in dados.columns]]
    if "vencimento" in dados:
        dados["vencimento"] = dados["vencimento"].apply(formatar_data_br)
    if "valor" in dados:
        dados["valor"] = dados["valor"].apply(formatar_moeda_br)
    return dados.rename(
        columns={
            "vencimento": "Vencimento",
            "descricao": "Descrição",
            "fornecedor_cliente": "Fornecedor",
            "tipo_conta_nome": "Tipo de conta",
            "valor": "Valor",
            "status": "Status",
            "categoria": "Categoria",
            "documento": "Documento",
            "anexo_nome": "Anexo",
            "codigo_pagamento": "Codigo de barras / Pix",
            "criado_em": "Incluído em",
            "criado_por": "Incluído por",
        }
    )


def estilo_status_linha(linha: pd.Series) -> list[str]:
    status = str(linha.get("Status", "")).strip().lower()
    if conta_vencida(linha):
        estilo = "background-color: rgba(220, 38, 38, 0.10); color: #7f1d1d;"
    elif status in {"aberto", "pendente"}:
        estilo = "background-color: rgba(183, 121, 31, 0.08); color: #713f12;"
    else:
        estilo = ""
    return [estilo] * len(linha)


def exibir_contas(df: pd.DataFrame) -> None:
    tabela = preparar_tabela_contas(df)
    st.dataframe(
        tabela.style.apply(estilo_status_linha, axis=1),
        use_container_width=True,
        hide_index=True,
    )


ORDENACOES_CONTAS = {
    "vencimento": {"label": "Vencimento", "coluna": "vencimento_dt", "tipo": "data"},
    "descricao": {"label": "Descricao", "coluna": "descricao", "tipo": "texto"},
    "fornecedor": {"label": "Fornecedor", "coluna": "fornecedor_cliente", "tipo": "texto"},
    "valor": {"label": "Valor", "coluna": "valor", "tipo": "numero"},
    "status": {"label": "Status", "coluna": "status", "tipo": "texto"},
    "documento": {"label": "Documento", "coluna": "documento", "tipo": "texto"},
}


def selecionar_ordenacao_contas(campo: str) -> None:
    atual = st.session_state.get("contas_ordem_campo", "")
    crescente = st.session_state.get("contas_ordem_crescente", True)
    st.session_state.contas_ordem_campo = campo
    if campo == "vencimento":
        st.session_state.contas_ordem_crescente = False
    else:
        st.session_state.contas_ordem_crescente = not crescente if atual == campo else True


def rotulo_ordenacao(campo: str) -> str:
    info = ORDENACOES_CONTAS[campo]
    campo_atual = st.session_state.get("contas_ordem_campo", "vencimento")
    crescente = False if campo == "vencimento" else bool(st.session_state.get("contas_ordem_crescente", False))
    if campo_atual != campo:
        return f"{info['label']} ↕"
    return f"{info['label']} {'↑' if crescente else '↓'}"


def ordenar_contas_exibidas(contas: pd.DataFrame) -> pd.DataFrame:
    campo = st.session_state.get("contas_ordem_campo", "vencimento")
    if campo not in ORDENACOES_CONTAS:
        campo = "vencimento"

    info = ORDENACOES_CONTAS[campo]
    coluna = str(info["coluna"])
    crescente = False if campo == "vencimento" else bool(st.session_state.get("contas_ordem_crescente", True))
    dados = contas.copy()

    if info["tipo"] == "data":
        dados["_ordem"] = pd.to_datetime(dados.get(coluna, ""), errors="coerce")
    elif info["tipo"] == "numero":
        dados["_ordem"] = pd.to_numeric(dados.get(coluna, 0), errors="coerce").fillna(0)
    else:
        dados["_ordem"] = (
            dados.get(coluna, "")
            .fillna("")
            .astype(str)
            .map(lambda texto: unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").casefold())
        )

    return dados.sort_values("_ordem", ascending=crescente, na_position="last").drop(columns=["_ordem"], errors="ignore")


def exibir_contas_com_acoes(
    contas_exibidas: pd.DataFrame,
    df_base: pd.DataFrame,
    empresa: str,
    usuario: str,
    pode_alterar: bool,
) -> None:
    if contas_exibidas.empty:
        st.info("Nenhum registro encontrado para este filtro.")
        return

    contas = ordenar_contas_exibidas(contas_exibidas)
    larguras = [1.35, 2.4, 2.25, 1.25, 1, 1.35]
    if pode_alterar:
        larguras.append(1.45)
    cabecalho = st.columns(larguras)
    for col, campo in zip(cabecalho[:6], ORDENACOES_CONTAS.keys()):
        col.button(
            rotulo_ordenacao(campo),
            key=f"ordenar_contas_{campo}",
            on_click=selecionar_ordenacao_contas,
            args=(campo,),
            use_container_width=True,
        )
    if pode_alterar:
        cabecalho[6].caption("Ações")

    for indice, linha in contas.iterrows():
        nivel = nivel_prazo_conta(linha)
        indicador = indicador_prazo_conta(linha)
        cols = st.columns(larguras)
        escrever_celula_conta(cols[0], formatar_data_br(linha.get("vencimento")), nivel, nowrap=True, indicador=indicador)
        escrever_celula_conta(cols[1], str(linha.get("descricao", "")), nivel)
        escrever_celula_conta(cols[2], str(linha.get("fornecedor_cliente", "")), nivel)
        escrever_celula_conta(cols[3], formatar_moeda_br(float(linha.get("valor", 0) or 0)), nivel, nowrap=True)
        escrever_celula_conta(cols[4], "vencido" if nivel == "vencida" else str(linha.get("status", "")), nivel, nowrap=True, indicador=indicador)
        escrever_celula_conta(cols[5], str(linha.get("documento", "")), nivel)
        if not pode_alterar:
            continue
        acao_cols = cols[6].columns(3)
        if acao_cols[0].button("✏️", key=f"editar_{indice}", help="Editar conta"):
            selecionar_conta_para_edicao(indice)
            st.rerun()
        if acao_cols[1].button("✅", key=f"pagar_{indice}", help="Marcar como pago"):
            try:
                df_atualizado = marcar_conta_como_paga(df_base, indice, usuario)
                salvar_dados_empresa(normalizar_dataframe(df_atualizado), empresa)
                st.success("Conta marcada como paga.")
                st.rerun()
            except ValueError as erro:
                st.error(str(erro))
        if acao_cols[2].button("🗑️", key=f"excluir_{indice}", help="Excluir conta"):
            try:
                df_atualizado = excluir_conta_a_pagar(df_base, indice, usuario)
                salvar_dados_empresa(normalizar_dataframe(df_atualizado), empresa)
                st.success("Conta excluída da lista ativa.")
                st.rerun()
            except ValueError as erro:
                st.error(str(erro))


def limpar_nome_arquivo(nome: str) -> str:
    base = Path(nome).name
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", base).strip("._") or "anexo"


def salvar_anexo(uploaded_file: object, empresa: str, documento: str) -> tuple[str, str]:
    if uploaded_file is None:
        return "", ""

    ANEXOS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    empresa_slug = limpar_nome_arquivo(empresa).upper()
    doc_slug = limpar_nome_arquivo(documento or "documento")
    nome_final = f"{timestamp}_{empresa_slug}_{doc_slug}_{limpar_nome_arquivo(uploaded_file.name)}"
    caminho = ANEXOS_DIR / nome_final
    caminho.write_bytes(uploaded_file.getbuffer())
    return uploaded_file.name, str(caminho.relative_to(BASE_DIR))


def obter_tipo_conta(selecao: str) -> tuple[str, str]:
    texto = str(selecao or "").strip()
    if " - " not in texto:
        return "", texto

    codigo, nome = texto.split(" - ", 1)
    codigo = codigo.strip()
    return codigo, nome.strip() or TIPOS_CONTA_PAGAR.get(codigo, "")


def opcoes_tipo_conta_com_historico(df: pd.DataFrame) -> list[str]:
    opcoes = []
    if {"tipo_conta_codigo", "tipo_conta_nome"}.issubset(df.columns):
        for _, linha in df[["tipo_conta_codigo", "tipo_conta_nome"]].dropna(how="all").iterrows():
            codigo = str(linha.get("tipo_conta_codigo", "") or "").strip()
            nome = str(linha.get("tipo_conta_nome", "") or "").strip()
            if codigo and nome:
                opcoes.append(f"{codigo} - {nome}")
            elif nome:
                opcoes.append(nome)
            elif codigo:
                opcoes.append(codigo)

    opcoes.extend(f"{codigo} - {nome}" for codigo, nome in TIPOS_CONTA_PAGAR.items())
    unicas = []
    chaves_vistas = set()
    for opcao in opcoes:
        texto = str(opcao or "").strip()
        if not texto:
            continue
        chave = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").casefold()
        chave = " ".join(chave.split())
        if chave in OPCOES_OCULTAS:
            continue
        if chave in chaves_vistas:
            continue
        chaves_vistas.add(chave)
        unicas.append(texto)
    return unicas


def adicionar_conta_a_pagar(
    df: pd.DataFrame,
    empresa: str,
    usuario: str,
    dados_formulario: dict,
) -> pd.DataFrame:
    anexo_nome, anexo_caminho = salvar_anexo(
        dados_formulario["anexo"],
        empresa,
        dados_formulario["documento"],
    )
    tipo_codigo, tipo_nome = obter_tipo_conta(dados_formulario["tipo_conta"])
    nova_linha = {
        "empresa": empresa,
        "competencia": "",
        "tipo": "conta_a_pagar",
        "descricao": dados_formulario["descricao"],
        "fornecedor_cliente": dados_formulario["fornecedor"],
        "vencimento": dados_formulario["vencimento"].strftime("%Y-%m-%d"),
        "pagamento_recebimento": "",
        "valor": float(dados_formulario["valor"]),
        "status": dados_formulario["status"],
        "categoria": dados_formulario["categoria"],
        "observacao": dados_formulario["observacao"],
        "documento": dados_formulario["documento"],
        "tipo_conta_codigo": tipo_codigo,
        "tipo_conta_nome": tipo_nome,
        "anexo_nome": anexo_nome,
        "anexo_caminho": anexo_caminho,
        "codigo_pagamento": dados_formulario["codigo_pagamento"],
        "criado_em": agora_br(),
        "criado_por": usuario,
        "excluido_em": "",
        "excluido_por": "",
        "ativo": True,
    }
    return pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)


def excluir_conta_a_pagar(df: pd.DataFrame, indice: int, usuario: str) -> pd.DataFrame:
    df_atualizado = df.copy()
    if indice not in df_atualizado.index:
        raise ValueError("Conta selecionada não foi encontrada na base atual.")

    posicao = df_atualizado.index.get_loc(indice)
    if isinstance(posicao, slice):
        posicao = posicao.start
    elif not isinstance(posicao, int):
        posicoes = list(posicao)
        if not posicoes:
            raise ValueError("Conta selecionada não foi encontrada na base atual.")
        posicao = posicoes[0]

    df_atualizado.iat[posicao, df_atualizado.columns.get_loc("ativo")] = False
    df_atualizado.iat[posicao, df_atualizado.columns.get_loc("excluido_em")] = agora_br()
    df_atualizado.iat[posicao, df_atualizado.columns.get_loc("excluido_por")] = usuario
    return df_atualizado


def marcar_conta_como_paga(df: pd.DataFrame, indice: int, usuario: str) -> pd.DataFrame:
    df_atualizado = df.copy()
    if indice not in df_atualizado.index:
        raise ValueError("Conta selecionada não foi encontrada na base atual.")

    posicao = df_atualizado.index.get_loc(indice)
    if isinstance(posicao, slice):
        posicao = posicao.start
    elif not isinstance(posicao, int):
        posicoes = list(posicao)
        if not posicoes:
            raise ValueError("Conta selecionada não foi encontrada na base atual.")
        posicao = posicoes[0]

    df_atualizado.iat[posicao, df_atualizado.columns.get_loc("status")] = "pago"
    df_atualizado.iat[posicao, df_atualizado.columns.get_loc("pagamento_recebimento")] = datetime.now().strftime("%Y-%m-%d")
    df_atualizado.iat[posicao, df_atualizado.columns.get_loc("observacao")] = (
        str(df_atualizado.iat[posicao, df_atualizado.columns.get_loc("observacao")] or "").strip()
        + f" | Pago registrado por {usuario} em {agora_br()}"
    ).strip(" |")
    return df_atualizado


def editar_conta_a_pagar(df: pd.DataFrame, indice: int, usuario: str, dados_formulario: dict) -> pd.DataFrame:
    df_atualizado = df.copy()
    if indice not in df_atualizado.index:
        raise ValueError("Conta selecionada não foi encontrada na base atual.")

    posicao = df_atualizado.index.get_loc(indice)
    if isinstance(posicao, slice):
        posicao = posicao.start
    elif not isinstance(posicao, int):
        posicoes = list(posicao)
        if not posicoes:
            raise ValueError("Conta selecionada não foi encontrada na base atual.")
        posicao = posicoes[0]

    tipo_codigo, tipo_nome = obter_tipo_conta(dados_formulario["tipo_conta"])
    anexo_nome, anexo_caminho = salvar_anexo(
        dados_formulario.get("anexo"),
        str(df_atualizado.iat[posicao, df_atualizado.columns.get_loc("empresa")] or ""),
        dados_formulario["documento"],
    )
    atualizacoes = {
        "tipo": "conta_a_pagar",
        "descricao": dados_formulario["descricao"],
        "fornecedor_cliente": dados_formulario["fornecedor"],
        "vencimento": dados_formulario["vencimento"].strftime("%Y-%m-%d"),
        "valor": float(dados_formulario["valor"]),
        "status": dados_formulario["status"],
        "categoria": dados_formulario["categoria"],
        "observacao": dados_formulario["observacao"],
        "documento": dados_formulario["documento"],
        "tipo_conta_codigo": tipo_codigo,
        "tipo_conta_nome": tipo_nome,
        "codigo_pagamento": dados_formulario["codigo_pagamento"],
        "criado_por": str(df_atualizado.iat[posicao, df_atualizado.columns.get_loc("criado_por")] or usuario),
        "excluido_em": "",
        "excluido_por": "",
        "ativo": True,
    }
    if anexo_nome:
        atualizacoes["anexo_nome"] = anexo_nome
        atualizacoes["anexo_caminho"] = anexo_caminho
    for coluna, valor in atualizacoes.items():
        df_atualizado.iat[posicao, df_atualizado.columns.get_loc(coluna)] = valor
    return df_atualizado


def selecionar_conta_para_edicao(indice: int) -> None:
    st.session_state.conta_edicao_indice = indice


def selecionar_manutencao(secao: str) -> None:
    st.session_state.manutencao_ativa = secao


def formulario_inclusao(df: pd.DataFrame, empresa: str, usuario: str) -> None:
    st.markdown("### + Incluir conta a pagar")
    st.caption("A inclusao fica registrada com data, hora e usuario logado.")

    opcoes_tipo = opcoes_com_novo(opcoes_tipo_conta_com_historico(df))
    opcoes_descricao = opcoes_com_novo(opcoes_com_historico(df, "descricao", SUGESTOES_DESCRICAO))
    opcoes_fornecedor = opcoes_com_novo(opcoes_com_historico(df, "fornecedor_cliente", SUGESTOES_FORNECEDOR))

    c1, c2 = st.columns(2)
    descricao_sel, descricao_nova = campo_opcao_nova(
        c1,
        "Descricao",
        opcoes_descricao,
        "Digite a nova descricao",
        "inclusao_descricao",
    )
    fornecedor_sel, fornecedor_novo = campo_opcao_nova(
        c2,
        "Fornecedor",
        opcoes_fornecedor,
        "Digite o novo fornecedor",
        "inclusao_fornecedor",
    )
    tipo_conta_sel, tipo_conta_novo = campo_opcao_nova(
        st,
        "Tipo de conta",
        opcoes_tipo,
        "Digite o novo tipo de conta",
        "inclusao_tipo_conta",
    )

    with st.form("form_conta_a_pagar", clear_on_submit=True):
        c3, c4, c5 = st.columns([1, 1, 1])
        vencimento = c3.date_input("Vencimento", format="DD/MM/YYYY")
        valor_texto = c4.text_input("Valor", placeholder="Ex.: 1.000,00")
        status = c5.selectbox("Status", ["aberto", "pendente", "vencido"])

        observacao = st.text_area("Descricao / observacao livre", height=88)
        c6, c7 = st.columns([1, 1])
        anexo = c6.file_uploader("Anexo", type=["pdf", "png", "jpg", "jpeg", "xml", "csv", "xlsx"])
        codigo_pagamento = c7.text_area("Codigo de barras ou chave Pix", height=88)
        acao_cancelar, acao_salvar = st.columns([1, 2])
        cancelar = acao_cancelar.form_submit_button("Cancelar", use_container_width=True)
        enviar = acao_salvar.form_submit_button("Salvar conta a pagar", use_container_width=True)

    if cancelar:
        st.session_state.manutencao_ativa = ""
        st.rerun()

    if not enviar:
        return

    descricao = resolver_opcao_digitavel(descricao_sel, descricao_nova)
    fornecedor = resolver_opcao_digitavel(fornecedor_sel, fornecedor_novo)
    tipo_conta = resolver_opcao_digitavel(tipo_conta_sel, tipo_conta_novo)
    _, tipo_nome = obter_tipo_conta(tipo_conta)
    valor = parse_valor_br(valor_texto)

    if not descricao.strip() or not fornecedor.strip() or not tipo_conta.strip() or valor is None or valor <= 0:
        st.error("Preencha descricao, fornecedor, tipo de conta e valor maior que zero. Use ponto para milhar e virgula para centavos.")
        return
    documento_final = f"AP-{empresa}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    dados_formulario = {
        "descricao": descricao.strip(),
        "fornecedor": fornecedor.strip(),
        "vencimento": vencimento,
        "valor": valor,
        "status": status,
        "tipo_conta": tipo_conta,
        "categoria": tipo_nome.strip(),
        "documento": documento_final,
        "observacao": observacao.strip(),
        "anexo": anexo,
        "codigo_pagamento": codigo_pagamento.strip(),
    }
    df_atualizado = adicionar_conta_a_pagar(df, empresa, usuario, dados_formulario)
    salvar_dados_empresa(df_atualizado, empresa)
    st.success("Conta a pagar incluida com sucesso.")
    st.rerun()

def importacao_massa_contas(df: pd.DataFrame, empresa: str, usuario: str) -> None:
    st.markdown("### Importar contas por CSV")
    st.caption("Baixe o modelo, preencha as linhas e importe o arquivo CSV nesta empresa.")
    st.download_button(
        "Baixar modelo CSV",
        data=gerar_modelo_importacao_csv(),
        file_name="modelo_contas_a_pagar.csv",
        mime="text/csv",
        use_container_width=True,
    )

    with st.form("form_importacao_massa", clear_on_submit=True):
        arquivo = st.file_uploader("Arquivo CSV preenchido", type=["csv"])
        importar = st.form_submit_button("Importar contas do CSV", use_container_width=True)

    if not importar:
        return

    if arquivo is None:
        st.error("Selecione um arquivo CSV para importar.")
        return

    try:
        df_importado = ler_csv_importacao(arquivo)
        novas_contas, erros = preparar_importacao_contas(df_importado, empresa, usuario)
    except ValueError as erro:
        st.error(str(erro))
        return

    if erros:
        st.warning("Algumas linhas não foram importadas.")
        for erro in erros[:10]:
            st.write(f"- {erro}")
        if len(erros) > 10:
            st.write(f"- Mais {len(erros) - 10} erro(s).")

    if novas_contas.empty:
        st.error("Nenhuma conta válida encontrada para importar.")
        return

    df_atualizado = pd.concat([df, novas_contas], ignore_index=True)
    salvar_dados_empresa(df_atualizado, empresa)
    st.success(f"{len(novas_contas)} conta(s) importada(s) com sucesso.")
    st.rerun()


def formulario_edicao_conta(df: pd.DataFrame, indice: int, empresa: str, usuario: str, key_prefix: str) -> None:
    if indice not in df.index:
        st.error("Conta selecionada nao foi encontrada.")
        if st.button("Fechar", key=f"{key_prefix}_fechar_indice_invalido"):
            st.session_state.pop("conta_edicao_indice", None)
            st.rerun()
        return

    linha = df.loc[indice]
    vencimento_atual = pd.to_datetime(linha.get("vencimento"), errors="coerce")
    if pd.isna(vencimento_atual):
        vencimento_atual = pd.Timestamp(datetime.now().date())

    opcoes_tipo = opcoes_com_novo(opcoes_tipo_conta_com_historico(df))
    tipo_codigo_atual = str(linha.get("tipo_conta_codigo", "") or "").strip()
    tipo_nome_atual = str(linha.get("tipo_conta_nome", "") or "").strip()
    tipo_rotulo = f"{tipo_codigo_atual} - {tipo_nome_atual}" if tipo_codigo_atual and tipo_nome_atual else tipo_nome_atual
    if not tipo_rotulo:
        tipo_rotulo = opcoes_tipo[indice_padrao_sem_novo(opcoes_tipo)]
    if tipo_rotulo not in opcoes_tipo:
        opcoes_tipo.insert(0, tipo_rotulo)
    tipo_idx = opcoes_tipo.index(tipo_rotulo)

    st.caption(
        f"{formatar_data_br(linha.get('vencimento'))} | "
        f"{linha.get('fornecedor_cliente', '')} | "
        f"{formatar_moeda_br(float(linha.get('valor', 0) or 0))}"
    )

    tipo_conta_sel, tipo_conta_novo = campo_opcao_nova(
        st,
        "Tipo de conta",
        opcoes_tipo,
        "Digite o novo tipo de conta",
        f"{key_prefix}_tipo_conta",
        index=tipo_idx,
    )

    with st.form(f"{key_prefix}_form_editar_conta"):
        c1, c2 = st.columns(2)
        descricao = c1.text_input("Descricao", value=str(linha.get("descricao", "")))
        fornecedor = c2.text_input("Fornecedor", value=str(linha.get("fornecedor_cliente", "")))

        c3, c4, c5 = st.columns([1, 1, 1])
        vencimento = c3.date_input("Vencimento", value=vencimento_atual.date(), format="DD/MM/YYYY")
        valor_texto = c4.text_input(
            "Valor",
            value=formatar_moeda_br(float(linha.get("valor", 0) or 0)).replace("R$ ", ""),
            placeholder="Ex.: 1.000,00",
        )
        status_opcoes = ["aberto", "pendente", "vencido"]
        status_atual = str(linha.get("status", "aberto")).lower()
        status = c5.selectbox("Status", status_opcoes, index=status_opcoes.index(status_atual) if status_atual in status_opcoes else 0)

        observacao = st.text_area(
            "Descricao / observacao livre",
            value=str(linha.get("observacao", "") or "").strip(),
            height=88,
        )
        anexo_atual = str(linha.get("anexo_nome", "") or "").strip()
        if anexo_atual:
            st.caption(f"Anexo atual: {anexo_atual}")
        c6, c7 = st.columns([1, 1])
        anexo = c6.file_uploader("Novo anexo", type=["pdf", "png", "jpg", "jpeg", "xml", "csv", "xlsx"])
        codigo_pagamento = c7.text_area(
            "Codigo de barras ou chave Pix",
            value=str(linha.get("codigo_pagamento", "") or "").strip(),
            height=88,
        )
        b1, b2 = st.columns(2)
        salvar = b1.form_submit_button("Salvar alteracoes", use_container_width=True)
        cancelar = b2.form_submit_button("Cancelar", use_container_width=True)

    if cancelar:
        st.session_state.pop("conta_edicao_indice", None)
        st.rerun()
    if not salvar:
        return

    tipo_conta = resolver_opcao_digitavel(tipo_conta_sel, tipo_conta_novo)
    _, tipo_nome = obter_tipo_conta(tipo_conta)
    valor = parse_valor_br(valor_texto)
    if not descricao.strip() or not fornecedor.strip() or not tipo_conta.strip() or valor is None or valor <= 0:
        st.error("Preencha descricao, fornecedor, tipo de conta e valor maior que zero. Use ponto para milhar e virgula para centavos.")
        return
    dados_formulario = {
        "descricao": descricao.strip(),
        "fornecedor": fornecedor.strip(),
        "vencimento": vencimento,
        "valor": valor,
        "status": status,
        "tipo_conta": tipo_conta,
        "categoria": tipo_nome.strip(),
        "documento": str(linha.get("documento", "") or "").strip() or f"AP-{empresa}-{datetime.now().strftime('%Y%m%d%H%M%S')}",
        "observacao": observacao.strip(),
        "anexo": anexo,
        "codigo_pagamento": codigo_pagamento.strip(),
    }
    try:
        df_atualizado = editar_conta_a_pagar(df, indice, usuario, dados_formulario)
    except ValueError as erro:
        st.error(str(erro))
        return

    salvar_dados_empresa(normalizar_dataframe(df_atualizado), empresa)
    st.session_state.pop("conta_edicao_indice", None)
    st.success("Conta a pagar atualizada com sucesso.")
    st.rerun()


if hasattr(st, "dialog"):
    @st.dialog("Editar conta a pagar")
    def janela_edicao_conta(df: pd.DataFrame, indice: int, empresa: str, usuario: str) -> None:
        formulario_edicao_conta(df, indice, empresa, usuario, "janela_edicao")
else:
    def janela_edicao_conta(df: pd.DataFrame, indice: int, empresa: str, usuario: str) -> None:
        st.warning("Atualize o Streamlit para abrir edicao em janela. Usando edicao na pagina.")
        formulario_edicao_conta(df, indice, empresa, usuario, "janela_edicao")


def area_edicao(df: pd.DataFrame, contas: pd.DataFrame, empresa: str, usuario: str) -> None:
    st.markdown("### Editar conta a pagar")
    st.caption("Selecione uma conta ativa, altere os campos e salve.")

    if contas.empty:
        st.info("Nenhuma conta em aberto para editar.")
        return

    opcoes = {}
    for indice, linha in contas.iterrows():
        rotulo = (
            f"{linha.get('documento', '')} | {formatar_data_br(linha.get('vencimento'))} | "
            f"{linha.get('fornecedor_cliente', '')} | {formatar_moeda_br(linha.get('valor', 0))}"
        )
        opcoes[rotulo] = indice

    labels = list(opcoes.keys())
    selecionada = st.selectbox("Conta a editar", labels, key="editar_conta_select")
    formulario_edicao_conta(df, opcoes[selecionada], empresa, usuario, "area_edicao")


def area_exclusao(df: pd.DataFrame, contas: pd.DataFrame, empresa: str, usuario: str) -> None:
    st.markdown("### 🗑️ Excluir conta a pagar")
    st.caption("A exclusão mantém auditoria no arquivo, marcando o registro como inativo.")

    if contas.empty:
        st.info("Nenhuma conta em aberto para excluir neste filtro.")
        return

    opcoes = {}
    for indice, linha in contas.iterrows():
        rotulo = (
            f"{linha.get('documento', '')} | {formatar_data_br(linha.get('vencimento'))} | "
            f"{linha.get('fornecedor_cliente', '')} | {formatar_moeda_br(linha.get('valor', 0))}"
        )
        opcoes[rotulo] = indice

    with st.form("form_excluir_conta"):
        selecionada = st.selectbox("Conta a excluir", list(opcoes.keys()))
        confirmar = st.checkbox("Confirmo a exclusão desta conta a pagar")
        excluir = st.form_submit_button("Excluir conta selecionada", use_container_width=True)

    if excluir:
        if not confirmar:
            st.warning("Marque a confirmação antes de excluir.")
            return

        df_atualizado = excluir_conta_a_pagar(df, opcoes[selecionada], usuario)
        salvar_dados_empresa(normalizar_dataframe(df_atualizado), empresa)
        st.success("Conta a pagar excluída da lista ativa.")
        st.rerun()


def pagina_contas_a_pagar(df: pd.DataFrame, empresa: str, usuario: str, config: dict) -> None:
    contas = filtrar_contas_a_pagar_abertas(df, empresa)
    pode_alterar = usuario_pode_alterar(config, usuario)
    hoje = pd.Timestamp(datetime.now().date())
    total_aberto = contas["valor"].sum()
    vencidas = contas.loc[contas.apply(conta_vencida, axis=1)]
    vence_hoje = contas.loc[contas["vencimento_dt"].dt.date == hoje.date()]
    proximos_7 = contas.loc[
        (contas["vencimento_dt"] > hoje)
        & (contas["vencimento_dt"] <= hoje + pd.Timedelta(days=7))
    ]
    abertas_no_prazo = contas.loc[
        (contas["status"].astype(str).str.lower().isin(["aberto", "pendente"]))
        & (contas["vencimento_dt"] > hoje + pd.Timedelta(days=7))
    ]

    st.subheader("Painel de pagamentos")
    st.caption("Visao para acompanhar o que precisa ser pago por vencimento.")

    renderizar_metricas(
        [
            {"label": "Total a pagar", "value": formatar_moeda_br(total_aberto)},
            {"label": "Vencidas", "value": formatar_moeda_br(vencidas["valor"].sum()), "delta": int(len(vencidas))},
            {"label": "Vence hoje", "value": formatar_moeda_br(vence_hoje["valor"].sum()), "delta": int(len(vence_hoje))},
            {"label": "Proximos 7 dias", "value": formatar_moeda_br(proximos_7["valor"].sum()), "delta": int(len(proximos_7))},
            {"label": "Contas em aberto", "value": int(len(contas))},
            {"label": "No prazo", "value": int(len(abertas_no_prazo))},
            {
                "label": "Ticket medio",
                "value": formatar_moeda_br(total_aberto / len(contas)) if len(contas) else formatar_moeda_br(0),
            },
        ]
    )

    if not len(contas):
        st.success("Nenhuma conta a pagar em aberto para esta empresa.")

    st.markdown("### Contas para pagar")
    st.download_button(
        "Baixar contas em Excel",
        data=gerar_excel_contas_download(contas, empresa),
        file_name=f"contas_a_pagar_{slug_empresa(empresa)}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    if contas.empty:
        st.info("Nenhum registro encontrado para este filtro.")
    else:
        exibir_contas_com_acoes(contas, df, empresa, usuario, pode_alterar)

    indice_edicao = st.session_state.get("conta_edicao_indice")
    if indice_edicao is not None and pode_alterar:
        janela_edicao_conta(df, indice_edicao, empresa, usuario)

    if not pode_alterar:
        st.info("Perfil de consulta: este usuario pode visualizar e baixar relatórios, mas nao pode incluir, editar, pagar, importar ou excluir contas.")
        st.session_state.manutencao_ativa = ""
        st.session_state.pop("conta_edicao_indice", None)
        return

    st.divider()
    st.markdown("### Manutencao")

    st.session_state.setdefault("manutencao_ativa", "")
    a1, a2, a3, espaco = st.columns([1.2, 1.2, 1.2, 5.4], gap="small")
    a1.button("➕ Incluir", help="Incluir nova conta a pagar", key="btn_manutencao_incluir", on_click=selecionar_manutencao, args=("incluir",), use_container_width=True)
    a2.button("📥 Importar", help="Importar contas em massa por CSV", key="btn_manutencao_importar", on_click=selecionar_manutencao, args=("importar",), use_container_width=True)
    a3.button("🗑️ Excluir", help="Excluir conta a pagar", key="btn_manutencao_excluir", on_click=selecionar_manutencao, args=("excluir",), use_container_width=True)

    if st.session_state.manutencao_ativa == "incluir":
        st.markdown("#### ➕ Incluir nova conta a pagar")
        formulario_inclusao(df, empresa, usuario)
    elif st.session_state.manutencao_ativa == "importar":
        st.markdown("#### 📥 Importar contas em massa por CSV")
        importacao_massa_contas(df, empresa, usuario)
    elif st.session_state.manutencao_ativa == "excluir":
        st.markdown("#### 🗑️ Excluir conta a pagar")
        area_exclusao(df, contas, empresa, usuario)

def main() -> None:
    configurar_pagina()
    config = carregar_config()
    inicializar_sessao()

    if not st.session_state.autenticado:
        tela_login(config)
        st.stop()

    if not st.session_state.get("empresa_logada"):
        dashboard_empresas(config, carregar_dados_empresas(config))
        st.stop()

    empresa, usuario = sidebar_filtros(config)
    df = carregar_dados_empresa(empresa)
    contas = filtrar_contas_a_pagar_abertas(df, empresa)
    status_geral, status_tipo = status_geral_contas(contas)

    mostrar_cabecalho(empresa, status_geral, status_tipo)
    if st.session_state.get("modulo_atual", "contas_a_pagar") == "contas_a_pagar":
        pagina_contas_a_pagar(df, empresa, usuario, config)
    else:
        st.info("Este módulo ainda não está liberado.")


if __name__ == "__main__":
    main()
