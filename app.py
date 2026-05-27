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
                <div class="meta-pill"><span>Empresa</span>{escape(empresa)}</div>
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
                "empresa": "Empresa Modelo LTDA",
                "senha": "demo123",
                "competencias": ["2026-01", "2026-02", "2026-03"],
            },
            {
                "empresa": "Comercio Exemplo SA",
                "senha": "cliente456",
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
        ["Empresa Modelo LTDA", "2026-01", "conta_a_pagar", "Aluguel da sede", "Imobiliaria Central", "2026-01-10", "", 3200.00, "aberto", "Administrativo", "Pagamento recorrente mensal", "AP-2026-001"],
        ["Empresa Modelo LTDA", "2026-01", "conta_a_pagar", "Energia eletrica", "Companhia de Energia", "2026-01-18", "", 860.45, "vencido", "Despesas fixas", "Conferir segunda via", "AP-2026-002"],
        ["Empresa Modelo LTDA", "2026-01", "conta_paga", "Internet corporativa", "Telecom Brasil", "2026-01-08", "2026-01-07", 349.90, "pago", "Despesas fixas", "Pago via debito automatico", "PG-2026-001"],
        ["Empresa Modelo LTDA", "2026-01", "receita", "Servicos de consultoria", "Cliente Alfa", "2026-01-15", "2026-01-15", 12500.00, "recebido", "Receita operacional", "Nota fiscal emitida", "NF-1001"],
        ["Empresa Modelo LTDA", "2026-01", "imposto", "DAS Simples Nacional", "Receita Federal", "2026-02-20", "", 1480.00, "pendente", "Tributos", "Guia em preparacao", "IMP-2026-001"],
        ["Empresa Modelo LTDA", "2026-01", "relatorio", "Balancete mensal", "Escritorio Contabil", "2026-01-31", "", 0.00, "disponível", "Relatorios contabeis", "PDF disponivel no portal", "REL-2026-001"],
        ["Empresa Modelo LTDA", "2026-02", "conta_a_pagar", "Licenca de software", "SaaS Financeiro", "2026-02-12", "", 590.00, "aberto", "Tecnologia", "Renovacao mensal", "AP-2026-003"],
        ["Empresa Modelo LTDA", "2026-02", "conta_paga", "Honorarios contabeis", "Escritorio Contabil", "2026-02-05", "2026-02-05", 980.00, "pago", "Contabilidade", "Pagamento confirmado", "PG-2026-002"],
        ["Empresa Modelo LTDA", "2026-02", "receita", "Mensalidade contrato Beta", "Cliente Beta", "2026-02-20", "2026-02-19", 8700.00, "recebido", "Receita recorrente", "Recebido por PIX", "NF-1035"],
        ["Empresa Modelo LTDA", "2026-02", "imposto", "INSS folha", "Receita Federal", "2026-03-20", "", 2100.00, "pendente", "Folha", "Valor estimado", "IMP-2026-002"],
        ["Empresa Modelo LTDA", "2026-02", "relatorio", "DRE gerencial", "Escritorio Contabil", "2026-02-28", "", 0.00, "disponível", "Relatorios gerenciais", "Demonstrativo para reuniao mensal", "REL-2026-002"],
        ["Comercio Exemplo SA", "2026-01", "conta_a_pagar", "Compra de mercadorias", "Distribuidora Prime", "2026-01-22", "", 6900.00, "aberto", "Estoque", "Boleto aguardando aprovacao", "AP-2026-101"],
        ["Comercio Exemplo SA", "2026-01", "conta_paga", "Frete de entregas", "Transportes Rapidos", "2026-01-11", "2026-01-11", 780.00, "pago", "Logistica", "Pago por transferencia", "PG-2026-101"],
        ["Comercio Exemplo SA", "2026-01", "receita", "Vendas no varejo", "Clientes diversos", "2026-01-31", "2026-01-31", 22400.00, "recebido", "Vendas", "Consolidado demonstrativo", "REC-2026-101"],
        ["Comercio Exemplo SA", "2026-01", "imposto", "ICMS apuracao", "Secretaria da Fazenda", "2026-02-12", "", 3260.00, "pendente", "Tributos estaduais", "Aguardando guia", "IMP-2026-101"],
        ["Comercio Exemplo SA", "2026-01", "relatorio", "Resumo fiscal", "Escritorio Contabil", "2026-01-31", "", 0.00, "disponível", "Relatorios fiscais", "Arquivo demonstrativo", "REL-2026-101"],
        ["Comercio Exemplo SA", "2026-02", "conta_a_pagar", "Manutencao loja", "Servicos Prediais", "2026-02-16", "", 1450.00, "aberto", "Operacional", "Servico aprovado", "AP-2026-102"],
        ["Comercio Exemplo SA", "2026-02", "conta_paga", "Sistema de vendas", "PDV Cloud", "2026-02-09", "2026-02-08", 420.00, "pago", "Tecnologia", "Assinatura mensal", "PG-2026-102"],
        ["Comercio Exemplo SA", "2026-02", "receita", "Vendas online", "Marketplace", "2026-02-28", "2026-02-28", 18100.00, "recebido", "E-commerce", "Repasse liquido demonstrativo", "REC-2026-102"],
        ["Comercio Exemplo SA", "2026-02", "imposto", "PIS/COFINS", "Receita Federal", "2026-03-25", "", 1980.00, "pendente", "Tributos federais", "Valor sujeito a revisao", "IMP-2026-102"],
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
    cliente = next((item for item in config["clientes"] if item["empresa"] == empresa), None)
    return bool(cliente and senha_digitada and senha_digitada == cliente["senha"])


def sidebar_filtros(config: dict) -> tuple[str, str, str]:
    if LOGO_HEADER_PATH.exists():
        st.sidebar.image(str(LOGO_HEADER_PATH), use_container_width=True)
    st.sidebar.title("🔐 Acesso do Cliente")
    st.sidebar.caption("Use apenas dados demonstrativos nesta versao.")

    empresas = [cliente["empresa"] for cliente in config["clientes"]]
    empresa = st.sidebar.selectbox("Empresa", empresas)

    cliente_atual = next(cliente for cliente in config["clientes"] if cliente["empresa"] == empresa)
    competencia = st.sidebar.selectbox("Competencia", cliente_atual["competencias"])

    senha = st.sidebar.text_input("Senha demonstrativa", type="password")
    st.sidebar.divider()
    st.sidebar.info("Senhas demo: `demo123` ou `cliente456` conforme a empresa selecionada.")
    return empresa, competencia, senha


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

    empresa, competencia, senha = sidebar_filtros(config)

    if not validar_login(config, empresa, senha):
        mostrar_cabecalho(empresa, competencia, "Aguardando senha", "pendente")
        st.divider()
        st.warning("Informe a senha demonstrativa correta na barra lateral para acessar os dados.")
        st.stop()

    df = carregar_dados()
    df_filtrado = filtrar_dados(df, empresa, competencia)
    status_geral, status_tipo = status_geral_competencia(df_filtrado)

    mostrar_cabecalho(empresa, competencia, status_geral, status_tipo)

    st.sidebar.success("Acesso liberado")

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
