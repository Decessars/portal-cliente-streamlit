from __future__ import annotations

import base64
import json
import re
from datetime import datetime
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
ASSETS_DIR = BASE_DIR / "assets"
ANEXOS_DIR = DATA_DIR / "anexos"
PORTAL_CSV_PATH = DATA_DIR / "dados_portal.csv"
CSV_PATH = DATA_DIR / "dados_demo.csv"
XLSX_PATH = DATA_DIR / "dados_demo.xlsx"
CONFIG_PATH = BASE_DIR / "config_clientes.json"
LOGO_HEADER_PATH = ASSETS_DIR / "mh_log_logo_app_header.png"

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
    "criado_em",
    "criado_por",
    "excluido_em",
    "excluido_por",
    "ativo",
]

COLUNAS_DADOS = COLUNAS_BASE + COLUNAS_EXTRAS


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
            .login-logo {{
                display: flex;
                justify-content: center;
                margin-bottom: 0.75rem;
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
            div.stButton > button,
            div.stDownloadButton > button,
            div.stFormSubmitButton > button {{
                border-radius: 8px;
                border: 1px solid var(--mh-accent);
                background: var(--mh-accent);
                color: #ffffff;
                font-weight: 700;
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
                .portal-header {{
                    align-items: flex-start;
                    flex-direction: column;
                }}
                .portal-brand {{
                    align-items: flex-start;
                    flex-direction: column;
                }}
                .portal-meta {{
                    justify-content: flex-start;
                    min-width: 0;
                }}
                .portal-title {{
                    font-size: 1.55rem;
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


def mostrar_cabecalho(empresa: str, competencia: str, status_geral: str, status_tipo: str) -> None:
    logo_b64 = imagem_base64(LOGO_HEADER_PATH)
    logo_html = (
        f'<img src="data:image/png;base64,{logo_b64}" alt="MH LOG" style="width:154px;max-width:100%;height:auto;">'
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
                <div class="meta-pill"><span>Competência</span>{escape(competencia)}</div>
                <div class="meta-pill {status_classe}"><span>Status</span>{escape(status_geral)}</div>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def criar_config_demo() -> dict:
    config = {
        "clientes": [
            {"empresa": "DMLIMA", "senha": "123456", "competencias": ["2026-01", "2026-02", "2026-03"]},
            {"empresa": "VICTOR", "senha": "123456", "competencias": ["2026-01", "2026-02", "2026-03"]},
            {"empresa": "MHLOG", "senha": "123456", "competencias": ["2026-01", "2026-02", "2026-03"]},
            {"empresa": "MH BRASIL", "senha": "123456", "competencias": ["2026-01", "2026-02", "2026-03"]},
        ]
    }
    CONFIG_PATH.write_text(json.dumps(config, indent=2, ensure_ascii=False), encoding="utf-8")
    return config


def carregar_config() -> dict:
    if not CONFIG_PATH.exists():
        return criar_config_demo()

    with CONFIG_PATH.open("r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


def dados_demo() -> pd.DataFrame:
    registros = [
        ["DMLIMA", "2026-01", "conta_a_pagar", "Aluguel da sede", "Imobiliaria Central", "2026-01-10", "", 3200.00, "aberto", "Administrativo", "Pagamento recorrente mensal", "AP-DML-001", "2.1.5.01.001", "SALARIOS E ORDENADOS A PAGAR", "", "", "", "", "", "", True],
        ["DMLIMA", "2026-01", "conta_a_pagar", "DAS Simples Nacional", "Receita Federal", "2026-01-20", "", 1480.00, "pendente", "Tributos", "Guia em preparacao", "IMP-DML-001", "2.1.4.01.015", "SIMPLES NACIONAL A RECOLHER", "", "", "", "", "", "", True],
        ["DMLIMA", "2026-02", "conta_a_pagar", "Licenca de software", "SaaS Financeiro", "2026-02-12", "", 590.00, "aberto", "Tecnologia", "Renovacao mensal", "AP-DML-002", "", "", "", "", "", "", "", "", True],
        ["VICTOR", "2026-01", "conta_a_pagar", "Compra de mercadorias", "Distribuidora Prime", "2026-01-22", "", 6900.00, "aberto", "Estoque", "Boleto aguardando aprovacao", "AP-VIC-001", "", "", "", "", "", "", "", "", True],
        ["VICTOR", "2026-02", "conta_a_pagar", "Manutencao loja", "Servicos Prediais", "2026-02-16", "", 1450.00, "aberto", "Operacional", "Servico aprovado", "AP-VIC-002", "", "", "", "", "", "", "", "", True],
        ["MHLOG", "2026-01", "conta_a_pagar", "Energia eletrica", "Companhia de Energia", "2026-01-18", "", 860.45, "vencido", "Despesas fixas", "Conferir segunda via", "AP-MHL-001", "", "", "", "", "", "", "", "", True],
        ["MHLOG", "2026-01", "conta_a_pagar", "Honorarios contabeis", "Escritorio Contabil", "2026-01-05", "", 980.00, "aberto", "Contabilidade", "Pagamento previsto", "AP-MHL-002", "2.1.6.02.001", "HONORARIOS CONTABEIS", "", "", "", "", "", "", True],
        ["MHLOG", "2026-02", "conta_a_pagar", "Manutencao de frota", "Oficina Parceira", "2026-02-19", "", 2250.00, "aberto", "Operacional", "Servico aprovado", "AP-MHL-003", "", "", "", "", "", "", "", "", True],
        ["MH BRASIL", "2026-01", "conta_a_pagar", "Servicos administrativos", "Fornecedor Brasil", "2026-01-14", "", 1850.00, "aberto", "Administrativo", "Boleto aguardando aprovacao", "AP-MHB-001", "", "", "", "", "", "", "", "", True],
        ["MH BRASIL", "2026-02", "conta_a_pagar", "Consultoria operacional", "Consultoria Beta", "2026-02-18", "", 2750.00, "aberto", "Operacional", "Servico programado", "AP-MHB-002", "", "", "", "", "", "", "", "", True],
    ]
    return pd.DataFrame(registros, columns=COLUNAS_DADOS)


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


def carregar_dados() -> pd.DataFrame:
    """Carrega dados do portal, dados demo ou recria a base demonstrativa."""
    if PORTAL_CSV_PATH.exists():
        df = pd.read_csv(PORTAL_CSV_PATH)
    elif CSV_PATH.exists():
        df = pd.read_csv(CSV_PATH)
    elif XLSX_PATH.exists():
        df = pd.read_excel(XLSX_PATH)
    else:
        df = criar_dados_demo()

    for coluna in COLUNAS_DADOS:
        if coluna not in df.columns:
            df[coluna] = True if coluna == "ativo" else ""

    df = df[COLUNAS_DADOS].copy()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0)
    df["ativo"] = df["ativo"].apply(normalizar_ativo)
    df["vencimento_dt"] = pd.to_datetime(df["vencimento"], errors="coerce")
    df["pagamento_recebimento_dt"] = pd.to_datetime(df["pagamento_recebimento"], errors="coerce")
    return df


def salvar_dados(df: pd.DataFrame) -> Path:
    """Salva no CSV operacional do portal, preservando dados demo separados."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    dados = df[[coluna for coluna in COLUNAS_DADOS if coluna in df.columns]].copy()
    dados.to_csv(PORTAL_CSV_PATH, index=False, encoding="utf-8-sig")
    return PORTAL_CSV_PATH


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


def validar_login(config: dict, empresa: str, senha_digitada: str) -> bool:
    senha_configurada = obter_senha_cliente(config, empresa)
    return bool(senha_configurada and senha_digitada and senha_digitada == senha_configurada)


def obter_cliente(config: dict, empresa: str) -> dict:
    return next((item for item in config["clientes"] if item["empresa"] == empresa), {})


def obter_senha_cliente(config: dict, empresa: str) -> str:
    """Lê senha do Streamlit Secrets quando existir; senão usa o JSON demo local."""
    try:
        clientes_secret = st.secrets.get("clientes", {})
        if empresa in clientes_secret:
            return str(clientes_secret[empresa])
    except Exception:
        pass

    return str(obter_cliente(config, empresa).get("senha", ""))


def inicializar_sessao() -> None:
    st.session_state.setdefault("autenticado", False)
    st.session_state.setdefault("empresa_logada", "")
    st.session_state.setdefault("competencia_logada", "")


def tela_login(config: dict) -> None:
    empresas = [cliente["empresa"] for cliente in config["clientes"]]
    empresa_padrao = st.session_state.get("empresa_login", empresas[0] if empresas else "")
    empresa_idx = empresas.index(empresa_padrao) if empresa_padrao in empresas else 0

    logo_b64 = imagem_base64(LOGO_HEADER_PATH)
    logo_html = (
        f'<img src="data:image/png;base64,{logo_b64}" alt="MH LOG" style="width:154px;max-width:100%;height:auto;">'
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


def sidebar_filtros(config: dict) -> tuple[str, str]:
    if LOGO_HEADER_PATH.exists():
        st.sidebar.image(str(LOGO_HEADER_PATH), use_container_width=True)
    st.sidebar.title("🔐 Sessão do Cliente")

    empresa = st.session_state.get("empresa_logada", "")
    cliente_atual = obter_cliente(config, empresa)
    competencias = cliente_atual.get("competencias", [])
    competencia_atual = st.session_state.get("competencia_logada", "")
    competencia_idx = competencias.index(competencia_atual) if competencia_atual in competencias else 0

    st.sidebar.text_input("Usuário", value=empresa, disabled=True)
    competencia = st.sidebar.selectbox("Competência", competencias, index=competencia_idx)
    st.session_state.competencia_logada = competencia
    st.sidebar.divider()
    st.sidebar.success("Acesso liberado")
    st.sidebar.button("Sair", use_container_width=True, on_click=logout)
    return empresa, competencia


def filtrar_dados(df: pd.DataFrame, empresa: str, competencia: str) -> pd.DataFrame:
    filtro = (df["empresa"] == empresa) & (df["competencia"] == competencia)
    return df.loc[filtro].copy()


def filtrar_contas_a_pagar_abertas(df: pd.DataFrame, empresa: str, competencia: str) -> pd.DataFrame:
    dados = filtrar_dados(df, empresa, competencia)
    status_abertos = dados["status"].astype(str).str.lower().isin(["aberto", "pendente", "vencido"])
    filtro = (dados["tipo"] == "conta_a_pagar") & status_abertos & dados["ativo"]
    return dados.loc[filtro].copy()


def status_geral_contas(df: pd.DataFrame) -> tuple[str, str]:
    if df.empty:
        return "Sem contas em aberto", "ok"

    vencidos = int((df["status"].astype(str).str.lower() == "vencido").sum())
    if vencidos:
        return f"{vencidos} conta(s) vencida(s)", "alerta"
    return f"{len(df)} conta(s) em aberto", "pendente"


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
        "criado_em",
        "criado_por",
    ]
    dados = df.sort_values(["vencimento_dt", "descricao"], na_position="last").copy()
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
            "criado_em": "Incluído em",
            "criado_por": "Incluído por",
        }
    )


def estilo_status_linha(linha: pd.Series) -> list[str]:
    status = str(linha.get("Status", "")).strip().lower()
    if status == "vencido":
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
    codigo = selecao.split(" - ", 1)[0].strip()
    return codigo, TIPOS_CONTA_PAGAR.get(codigo, "")


def adicionar_conta_a_pagar(
    df: pd.DataFrame,
    empresa: str,
    competencia: str,
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
        "competencia": competencia,
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
        "criado_em": agora_br(),
        "criado_por": usuario,
        "excluido_em": "",
        "excluido_por": "",
        "ativo": True,
    }
    return pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)


def excluir_conta_a_pagar(df: pd.DataFrame, indice: int, usuario: str) -> pd.DataFrame:
    df_atualizado = df.copy()
    df_atualizado.loc[indice, "ativo"] = False
    df_atualizado.loc[indice, "excluido_em"] = agora_br()
    df_atualizado.loc[indice, "excluido_por"] = usuario
    return df_atualizado


def formulario_inclusao(df: pd.DataFrame, empresa: str, competencia: str, usuario: str) -> None:
    st.markdown("### ➕ Incluir conta a pagar")
    st.caption("A inclusão fica registrada com data, hora e usuário logado.")

    opcoes_tipo = [f"{codigo} - {nome}" for codigo, nome in TIPOS_CONTA_PAGAR.items()]
    with st.form("form_conta_a_pagar", clear_on_submit=True):
        c1, c2 = st.columns(2)
        descricao = c1.text_input("Descrição", placeholder="Ex.: Honorários contábeis")
        fornecedor = c2.text_input("Fornecedor", placeholder="Ex.: Escritório Contábil")

        c3, c4, c5 = st.columns([1, 1, 1])
        vencimento = c3.date_input("Vencimento")
        valor = c4.number_input("Valor", min_value=0.0, step=10.0, format="%.2f")
        status = c5.selectbox("Status", ["aberto", "pendente", "vencido"])

        tipo_conta = st.selectbox("Tipo de conta", opcoes_tipo)
        c6, c7 = st.columns(2)
        categoria = c6.text_input("Categoria", placeholder="Ex.: Tributos, Folha, Contabilidade")
        documento = c7.text_input("Documento / referência", placeholder="Ex.: NF-123, BOLETO-001")

        observacao = st.text_area("Observação", height=80)
        anexo = st.file_uploader(
            "Anexo",
            type=["pdf", "png", "jpg", "jpeg", "xlsx", "csv", "doc", "docx"],
            help="Use para anexar boleto, nota fiscal, guia ou comprovante.",
        )
        enviar = st.form_submit_button("Salvar conta a pagar", use_container_width=True)

    if not enviar:
        return

    if not descricao.strip() or not fornecedor.strip() or valor <= 0:
        st.error("Preencha descrição, fornecedor e valor maior que zero.")
        return

    documento_final = documento.strip() or f"AP-{empresa}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    dados_formulario = {
        "descricao": descricao.strip(),
        "fornecedor": fornecedor.strip(),
        "vencimento": vencimento,
        "valor": valor,
        "status": status,
        "tipo_conta": tipo_conta,
        "categoria": categoria.strip(),
        "documento": documento_final,
        "observacao": observacao.strip(),
        "anexo": anexo,
    }
    df_atualizado = adicionar_conta_a_pagar(df, empresa, competencia, usuario, dados_formulario)
    salvar_dados(df_atualizado)
    st.success("Conta a pagar incluída com sucesso.")
    st.rerun()


def area_exclusao(df: pd.DataFrame, contas: pd.DataFrame, usuario: str) -> None:
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
        salvar_dados(df_atualizado)
        st.success("Conta a pagar excluída da lista ativa.")
        st.rerun()


def pagina_contas_a_pagar(df: pd.DataFrame, empresa: str, competencia: str, usuario: str) -> None:
    contas = filtrar_contas_a_pagar_abertas(df, empresa, competencia)
    total_aberto = contas["valor"].sum()
    vencidas = contas.loc[contas["status"].astype(str).str.lower() == "vencido"]
    proximos = contas.sort_values(["vencimento_dt", "descricao"], na_position="last").head(5)

    st.subheader("💸 Contas a Pagar")
    st.caption("Lista operacional de obrigações em aberto, com inclusão manual e anexos.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total em aberto", formatar_moeda_br(total_aberto))
    c2.metric("Quantidade", int(len(contas)))
    c3.metric("Vencidas", int(len(vencidas)))
    c4.metric("Próximos vencimentos", int(len(proximos)))

    if len(vencidas):
        st.error("Existem contas vencidas neste filtro. Revise os vencimentos antes do pagamento.")
    elif len(contas):
        st.warning("Existem contas em aberto aguardando acompanhamento.")
    else:
        st.success("Nenhuma conta a pagar em aberto para esta competência.")

    with st.expander("➕ Incluir nova conta a pagar", expanded=False):
        formulario_inclusao(df, empresa, competencia, usuario)

    st.divider()
    st.markdown("### 📋 Contas em aberto")
    st.download_button(
        "⬇️ Baixar contas filtradas em CSV",
        data=gerar_csv_download(contas),
        file_name=f"contas_a_pagar_{empresa}_{competencia}.csv",
        mime="text/csv",
        use_container_width=True,
    )

    if contas.empty:
        st.info("Nenhum registro encontrado para este filtro.")
    else:
        exibir_contas(contas)

    st.divider()
    with st.expander("🗑️ Excluir conta a pagar", expanded=False):
        area_exclusao(df, contas, usuario)


def main() -> None:
    configurar_pagina()
    config = carregar_config()
    inicializar_sessao()

    if not st.session_state.autenticado:
        tela_login(config)
        st.stop()

    empresa, competencia = sidebar_filtros(config)
    df = carregar_dados()
    contas = filtrar_contas_a_pagar_abertas(df, empresa, competencia)
    status_geral, status_tipo = status_geral_contas(contas)

    mostrar_cabecalho(empresa, competencia, status_geral, status_tipo)
    pagina_contas_a_pagar(df, empresa, competencia, empresa)


if __name__ == "__main__":
    main()
