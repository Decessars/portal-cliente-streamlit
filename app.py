from __future__ import annotations

import base64
import json
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
ASSETS_DIR = BASE_DIR / "assets"
PORTAL_CSV_PATH = DATA_DIR / "dados_portal.csv"
CSV_PATH = DATA_DIR / "dados_demo.csv"
XLSX_PATH = DATA_DIR / "dados_demo.xlsx"
CONFIG_PATH = BASE_DIR / "config_clientes.json"
LOGO_HEADER_PATH = ASSETS_DIR / "mh_log_logo_app_header.png"
LOGO_ICON_PATH = ASSETS_DIR / "mh_log_logo_app_512.png"

TEMA_MH = {
    "name": "MH Verde Nevoa",
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
    "violet": "#4b8f65",
    "teal": "#1f9d68",
    "border": "#c8ddce",
}

COLUNAS_DADOS = [
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
            [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
            [data-testid="stSidebar"] label {{
                color: var(--mh-text);
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
            [data-testid="stMetricLabel"] {{
                color: var(--mh-muted);
            }}
            [data-testid="stMetricValue"] {{
                color: var(--mh-accent);
            }}
            .portal-card {{
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                padding: 16px;
                background: var(--mh-panel);
            }}
            .small-muted {{
                color: var(--mh-muted);
                font-size: 0.9rem;
            }}
            .info-card {{
                min-height: 11.5rem;
                border: 1px solid var(--mh-border);
                border-radius: 8px;
                padding: 1rem;
                background: var(--mh-panel);
                box-shadow: 0 6px 18px rgba(23, 51, 38, 0.07);
            }}
            .info-card h4 {{
                margin: 0 0 0.7rem;
                color: var(--mh-text);
                font-size: 1rem;
            }}
            .info-card p,
            .info-card li {{
                color: var(--mh-muted);
                font-size: 0.9rem;
                margin: 0.28rem 0;
            }}
            .info-card strong {{
                color: var(--mh-text);
            }}
            .card-ok {{
                border-left: 5px solid var(--mh-ok);
            }}
            .card-alerta {{
                border-left: 5px solid var(--mh-danger);
            }}
            .card-pendente {{
                border-left: 5px solid var(--mh-warning);
            }}
            .card-neutro {{
                border-left: 5px solid var(--mh-accent);
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
            div.stButton > button,
            div.stDownloadButton > button {{
                border-radius: 8px;
                border: 1px solid var(--mh-accent);
                background: var(--mh-accent);
                color: #ffffff;
                font-weight: 700;
            }}
            div.stButton > button:hover,
            div.stDownloadButton > button:hover {{
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
                        Sistema demonstrativo para acompanhamento contábil e financeiro.
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
    """Cria uma configuração demonstrativa se o JSON ainda não existir."""
    config = {
        "clientes": [
            {
                "empresa": "DMLIMA",
                "senha": "123456",
                "competencias": ["2026-01", "2026-02", "2026-03"],
            },
            {
                "empresa": "VICTOR",
                "senha": "123456",
                "competencias": ["2026-01", "2026-02", "2026-03"],
            },
            {
                "empresa": "MHLOG",
                "senha": "123456",
                "competencias": ["2026-01", "2026-02", "2026-03"],
            },
            {
                "empresa": "MH BRASIL",
                "senha": "123456",
                "competencias": ["2026-01", "2026-02", "2026-03"],
            },
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
        ["DMLIMA", "2026-01", "conta_a_pagar", "Aluguel da sede", "Imobiliaria Central", "2026-01-10", "", 3200.00, "aberto", "Administrativo", "Pagamento recorrente mensal", "AP-DML-001"],
        ["DMLIMA", "2026-01", "conta_paga", "Internet corporativa", "Telecom Brasil", "2026-01-08", "2026-01-07", 349.90, "pago", "Despesas fixas", "Pago via debito automatico", "PG-DML-001"],
        ["DMLIMA", "2026-01", "receita", "Servicos de consultoria", "Cliente Alfa", "2026-01-15", "2026-01-15", 12500.00, "recebido", "Receita operacional", "Nota fiscal emitida", "NF-DML-001"],
        ["DMLIMA", "2026-01", "imposto", "DAS Simples Nacional", "Receita Federal", "2026-02-20", "", 1480.00, "pendente", "Tributos", "Guia em preparacao", "IMP-DML-001"],
        ["DMLIMA", "2026-01", "relatorio", "Balancete mensal", "Escritorio Contabil", "2026-01-31", "", 0.00, "disponível", "Relatorios contabeis", "PDF disponivel no portal", "REL-DML-001"],
        ["DMLIMA", "2026-02", "conta_a_pagar", "Licenca de software", "SaaS Financeiro", "2026-02-12", "", 590.00, "aberto", "Tecnologia", "Renovacao mensal", "AP-DML-002"],
        ["DMLIMA", "2026-02", "receita", "Mensalidade contrato Beta", "Cliente Beta", "2026-02-20", "2026-02-19", 8700.00, "recebido", "Receita recorrente", "Recebido por PIX", "NF-DML-002"],
        ["VICTOR", "2026-01", "conta_a_pagar", "Compra de mercadorias", "Distribuidora Prime", "2026-01-22", "", 6900.00, "aberto", "Estoque", "Boleto aguardando aprovacao", "AP-VIC-001"],
        ["VICTOR", "2026-01", "conta_paga", "Frete de entregas", "Transportes Rapidos", "2026-01-11", "2026-01-11", 780.00, "pago", "Logistica", "Pago por transferencia", "PG-VIC-001"],
        ["VICTOR", "2026-01", "receita", "Vendas no varejo", "Clientes diversos", "2026-01-31", "2026-01-31", 22400.00, "recebido", "Vendas", "Consolidado demonstrativo", "REC-VIC-001"],
        ["VICTOR", "2026-01", "imposto", "ICMS apuracao", "Secretaria da Fazenda", "2026-02-12", "", 3260.00, "pendente", "Tributos estaduais", "Aguardando guia", "IMP-VIC-001"],
        ["VICTOR", "2026-01", "relatorio", "Resumo fiscal", "Escritorio Contabil", "2026-01-31", "", 0.00, "disponível", "Relatorios fiscais", "Arquivo demonstrativo", "REL-VIC-001"],
        ["VICTOR", "2026-02", "conta_a_pagar", "Manutencao loja", "Servicos Prediais", "2026-02-16", "", 1450.00, "aberto", "Operacional", "Servico aprovado", "AP-VIC-002"],
        ["VICTOR", "2026-02", "receita", "Vendas online", "Marketplace", "2026-02-28", "2026-02-28", 18100.00, "recebido", "E-commerce", "Repasse liquido demonstrativo", "REC-VIC-002"],
        ["MHLOG", "2026-01", "conta_a_pagar", "Energia eletrica", "Companhia de Energia", "2026-01-18", "", 860.45, "vencido", "Despesas fixas", "Conferir segunda via", "AP-MHL-001"],
        ["MHLOG", "2026-01", "conta_paga", "Honorarios contabeis", "Escritorio Contabil", "2026-01-05", "2026-01-05", 980.00, "pago", "Contabilidade", "Pagamento confirmado", "PG-MHL-001"],
        ["MHLOG", "2026-01", "receita", "Servicos logisticos", "Cliente Operacional", "2026-01-30", "2026-01-30", 28600.00, "recebido", "Receita operacional", "Consolidado demonstrativo", "REC-MHL-001"],
        ["MHLOG", "2026-01", "imposto", "PIS/COFINS", "Receita Federal", "2026-02-25", "", 1980.00, "pendente", "Tributos federais", "Valor sujeito a revisao", "IMP-MHL-001"],
        ["MHLOG", "2026-01", "relatorio", "DRE gerencial", "Escritorio Contabil", "2026-01-31", "", 0.00, "disponível", "Relatorios gerenciais", "Demonstrativo para reuniao mensal", "REL-MHL-001"],
        ["MHLOG", "2026-02", "conta_a_pagar", "Manutencao de frota", "Oficina Parceira", "2026-02-19", "", 2250.00, "aberto", "Operacional", "Servico aprovado", "AP-MHL-002"],
        ["MHLOG", "2026-02", "receita", "Contrato mensal transporte", "Cliente Logistico", "2026-02-28", "2026-02-28", 31400.00, "recebido", "Receita recorrente", "Recebido por transferencia", "REC-MHL-002"],
        ["MH BRASIL", "2026-01", "conta_a_pagar", "Servicos administrativos", "Fornecedor Brasil", "2026-01-14", "", 1850.00, "aberto", "Administrativo", "Boleto aguardando aprovacao", "AP-MHB-001"],
        ["MH BRASIL", "2026-01", "conta_paga", "Sistema financeiro", "SaaS Financeiro", "2026-01-09", "2026-01-09", 620.00, "pago", "Tecnologia", "Assinatura mensal", "PG-MHB-001"],
        ["MH BRASIL", "2026-01", "receita", "Contrato de apoio operacional", "Cliente Brasil", "2026-01-30", "2026-01-30", 19800.00, "recebido", "Receita operacional", "Recebimento demonstrativo", "REC-MHB-001"],
        ["MH BRASIL", "2026-01", "imposto", "DAS Simples Nacional", "Receita Federal", "2026-02-20", "", 2320.00, "pendente", "Tributos", "Guia demonstrativa em preparacao", "IMP-MHB-001"],
        ["MH BRASIL", "2026-01", "relatorio", "Resumo mensal", "Escritorio Contabil", "2026-01-31", "", 0.00, "disponível", "Relatorios contabeis", "Arquivo demonstrativo disponivel", "REL-MHB-001"],
        ["MH BRASIL", "2026-02", "conta_a_pagar", "Consultoria operacional", "Consultoria Beta", "2026-02-18", "", 2750.00, "aberto", "Operacional", "Servico programado", "AP-MHB-002"],
        ["MH BRASIL", "2026-02", "receita", "Mensalidade contrato Brasil", "Cliente Brasil", "2026-02-28", "2026-02-28", 21400.00, "recebido", "Receita recorrente", "Recebido por PIX", "REC-MHB-002"],
    ]
    return pd.DataFrame(registros, columns=COLUNAS_DADOS)


def criar_dados_demo() -> pd.DataFrame:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    df = dados_demo()
    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    return df


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
            df[coluna] = ""

    df = df[COLUNAS_DADOS].copy()
    df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0)
    df["vencimento_dt"] = pd.to_datetime(df["vencimento"], errors="coerce")
    df["pagamento_recebimento_dt"] = pd.to_datetime(df["pagamento_recebimento"], errors="coerce")
    return df


def formatar_moeda_br(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_data_br(valor: object) -> str:
    data = pd.to_datetime(valor, errors="coerce")
    if pd.isna(data):
        return ""
    return data.strftime("%d/%m/%Y")


def gerar_csv_download(df: pd.DataFrame) -> bytes:
    colunas = [coluna for coluna in COLUNAS_DADOS if coluna in df.columns]
    return df[colunas].to_csv(index=False).encode("utf-8-sig")


def formatar_moeda(valor: float) -> str:
    return formatar_moeda_br(valor)


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
                <p>Consulte informações contábeis e financeiras liberadas pelo escritório.</p>
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
                cliente_atual = obter_cliente(config, empresa)
                competencias = cliente_atual.get("competencias", [])
                senha = st.text_input("Senha de acesso", type="password")
                entrar = st.form_submit_button("Entrar", use_container_width=True)

            if entrar:
                if validar_login(config, empresa, senha):
                    competencia = competencias[0] if competencias else ""
                    st.session_state.autenticado = True
                    st.session_state.empresa_logada = empresa
                    st.session_state.competencia_logada = competencia
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


def dados_filtrados(df: pd.DataFrame, empresa: str, competencia: str) -> pd.DataFrame:
    return filtrar_dados(df, empresa, competencia)


def tabela_por_tipo(df: pd.DataFrame, tipo: str) -> pd.DataFrame:
    return df.loc[df["tipo"] == tipo].copy()


def mostrar_download(df: pd.DataFrame, nome_arquivo: str) -> None:
    st.download_button(
        "⬇️ Baixar dados filtrados em CSV",
        data=gerar_csv_download(df),
        file_name=nome_arquivo,
        mime="text/csv",
        use_container_width=True,
    )


def preparar_tabela(df: pd.DataFrame) -> pd.DataFrame:
    dados = df.sort_values(["vencimento_dt", "descricao"], na_position="last").copy()
    dados = dados[[coluna for coluna in COLUNAS_DADOS if coluna in dados.columns]]
    dados["vencimento"] = dados["vencimento"].apply(formatar_data_br)
    dados["pagamento_recebimento"] = dados["pagamento_recebimento"].apply(formatar_data_br)
    dados["valor"] = dados["valor"].apply(formatar_moeda_br)
    return dados


def estilo_status_linha(linha: pd.Series) -> list[str]:
    status = str(linha.get("status", "")).strip().lower()
    if status == "vencido":
        estilo = "background-color: rgba(220, 38, 38, 0.10); color: #7f1d1d;"
    elif status in {"pago", "recebido", "disponível"}:
        estilo = "background-color: rgba(21, 128, 61, 0.08); color: #14532d;"
    elif status in {"aberto", "pendente"}:
        estilo = "background-color: rgba(183, 121, 31, 0.08); color: #713f12;"
    else:
        estilo = ""
    return [estilo] * len(linha)


def exibir_dataframe(df: pd.DataFrame) -> None:
    tabela = preparar_tabela(df)
    st.dataframe(
        tabela.style.apply(estilo_status_linha, axis=1),
        use_container_width=True,
        hide_index=True,
    )


def mostrar_tabela(df: pd.DataFrame, tipo: str, nome_arquivo: str) -> None:
    dados = tabela_por_tipo(df, tipo).sort_values(["vencimento_dt", "descricao"], na_position="last")
    mostrar_download(dados, nome_arquivo)

    if dados.empty:
        st.warning("Nenhum registro encontrado para este filtro.")
        return

    exibir_dataframe(dados)


def status_geral_competencia(df: pd.DataFrame) -> tuple[str, str]:
    if df.empty:
        return "Sem dados", "pendente"

    vencidos = int((df["status"].astype(str).str.lower() == "vencido").sum())
    pendentes = int(df["status"].astype(str).str.lower().isin(["aberto", "pendente"]).sum())

    if vencidos:
        return f"{vencidos} item(ns) vencido(s)", "alerta"
    if pendentes:
        return f"{pendentes} item(ns) pendente(s)", "pendente"
    return "Em dia", "ok"


def card_resumo(titulo: str, linhas: list[str], classe: str = "card-neutro") -> None:
    itens = "".join(f"<li>{linha}</li>" for linha in linhas)
    st.markdown(
        f"""
        <div class="info-card {classe}">
            <h4>{titulo}</h4>
            <ul>{itens}</ul>
        </div>
        """,
        unsafe_allow_html=True,
    )


def aba_resumo(df: pd.DataFrame, empresa: str, competencia: str) -> None:
    total_a_pagar = df.loc[df["tipo"] == "conta_a_pagar", "valor"].sum()
    total_pago = df.loc[df["tipo"] == "conta_paga", "valor"].sum()
    total_recebido = df.loc[df["tipo"] == "receita", "valor"].sum()
    impostos_em_aberto = df.loc[df["tipo"] == "imposto", "valor"].sum()
    saldo_previsto = total_recebido - total_a_pagar - impostos_em_aberto

    st.subheader("🏠 Visao geral")
    st.caption(f"{empresa} | Competencia {competencia}")

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("💸 Total a pagar", formatar_moeda_br(total_a_pagar))
    col2.metric("✅ Total pago", formatar_moeda_br(total_pago))
    col3.metric("📥 Total recebido", formatar_moeda_br(total_recebido))
    col4.metric("🧾 Impostos em aberto", formatar_moeda_br(impostos_em_aberto))
    col5.metric("📊 Saldo previsto", formatar_moeda_br(saldo_previsto))

    st.divider()

    if saldo_previsto < 0:
        st.error("Atenção: o saldo financeiro previsto esta negativo para esta competencia.")
    elif impostos_em_aberto > 0:
        st.warning("Existem impostos pendentes. Confira os vencimentos antes do fechamento.")
    else:
        st.success("Fluxo financeiro demonstrativo sem alertas criticos.")

    vencimentos = df.loc[df["tipo"].isin(["conta_a_pagar", "imposto"])].copy()
    vencimentos = vencimentos.sort_values(["vencimento_dt", "descricao"], na_position="last").head(5)
    proximos = [
        f"<strong>{escape(formatar_data_br(linha.vencimento))}</strong> - {escape(str(linha.descricao))}: {escape(formatar_moeda_br(linha.valor))}"
        for linha in vencimentos.itertuples()
    ] or ["Nenhum vencimento cadastrado para a competência."]

    vencidos = int((df["status"].astype(str).str.lower() == "vencido").sum())
    pendentes = int(df["status"].astype(str).str.lower().isin(["aberto", "pendente"]).sum())
    docs_disponiveis = int((df["tipo"] == "relatorio").sum())
    impostos_pendentes = int(
        ((df["tipo"] == "imposto") & df["status"].astype(str).str.lower().isin(["aberto", "pendente"])).sum()
    )

    c1, c2 = st.columns(2)
    with c1:
        card_resumo(
            "Situação Financeira",
            [
                f"Saldo previsto: <strong>{escape(formatar_moeda_br(saldo_previsto))}</strong>",
                f"Recebimentos: <strong>{escape(formatar_moeda_br(total_recebido))}</strong>",
                f"Compromissos em aberto: <strong>{escape(formatar_moeda_br(total_a_pagar + impostos_em_aberto))}</strong>",
            ],
            "card-ok" if saldo_previsto >= 0 else "card-alerta",
        )
    with c2:
        card_resumo("Próximos Vencimentos", proximos, "card-pendente" if not vencimentos.empty else "card-neutro")

    c3, c4 = st.columns(2)
    with c3:
        card_resumo(
            "Alertas Contábeis",
            [
                f"Itens vencidos: <strong>{vencidos}</strong>",
                f"Itens abertos ou pendentes: <strong>{pendentes}</strong>",
                f"Impostos pendentes: <strong>{impostos_pendentes}</strong>",
            ],
            "card-alerta" if vencidos else "card-pendente" if pendentes else "card-ok",
        )
    with c4:
        card_resumo(
            "Documentos Disponíveis",
            [
                f"Relatórios cadastrados: <strong>{docs_disponiveis}</strong>",
                "Downloads em CSV disponíveis nas abas do portal.",
                "Área preparada para anexos PDF em versão futura.",
            ],
            "card-neutro",
        )

    st.divider()
    st.markdown("#### Movimentacoes da competencia")
    mostrar_download(df, f"dados_filtrados_{empresa}_{competencia}.csv")
    exibir_dataframe(df)


def main() -> None:
    configurar_pagina()
    config = carregar_config()
    inicializar_sessao()

    if not st.session_state.autenticado:
        tela_login(config)
        st.stop()

    empresa, competencia = sidebar_filtros(config)

    df = carregar_dados()
    df_filtrado = filtrar_dados(df, empresa, competencia)
    status_geral, status_tipo = status_geral_competencia(df_filtrado)

    mostrar_cabecalho(empresa, competencia, status_geral, status_tipo)

    if df_filtrado.empty:
        st.info("Nao ha dados demonstrativos para a empresa e competencia selecionadas.")
        st.stop()

    abas = st.tabs(
        [
            "🏠 Resumo",
            "💸 Contas a Pagar",
            "✅ Contas Pagas",
            "📥 Receitas",
            "🧾 Impostos",
            "📁 Relatórios",
        ]
    )

    with abas[0]:
        aba_resumo(df_filtrado, empresa, competencia)

    with abas[1]:
        st.subheader("💸 Contas a Pagar")
        mostrar_tabela(df_filtrado, "conta_a_pagar", f"contas_a_pagar_{empresa}_{competencia}.csv")

    with abas[2]:
        st.subheader("✅ Contas Pagas")
        mostrar_tabela(df_filtrado, "conta_paga", f"contas_pagas_{empresa}_{competencia}.csv")

    with abas[3]:
        st.subheader("📥 Receitas")
        mostrar_tabela(df_filtrado, "receita", f"receitas_{empresa}_{competencia}.csv")

    with abas[4]:
        st.subheader("🧾 Impostos")
        mostrar_tabela(df_filtrado, "imposto", f"impostos_{empresa}_{competencia}.csv")

    with abas[5]:
        st.subheader("📁 Relatórios")
        st.info("Nesta primeira versao, os relatorios aparecem como registros demonstrativos.")
        mostrar_tabela(df_filtrado, "relatorio", f"relatorios_{empresa}_{competencia}.csv")


if __name__ == "__main__":
    main()
