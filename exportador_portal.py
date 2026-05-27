from __future__ import annotations

from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).parent

COLUNAS_PORTAL = [
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

COLUNAS_OBRIGATORIAS = [
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

TIPOS_VALIDOS = {
    "conta_a_pagar",
    "conta_paga",
    "receita",
    "imposto",
    "relatorio",
}

STATUS_VALIDOS = {
    "aberto",
    "pago",
    "recebido",
    "vencido",
    "pendente",
    "disponível",
}


def normalizar_colunas_portal(df: pd.DataFrame) -> pd.DataFrame:
    """Garante as colunas esperadas pelo portal, na ordem correta."""
    dados = df.copy()
    dados.columns = [str(coluna).strip() for coluna in dados.columns]

    for coluna in COLUNAS_PORTAL:
        if coluna not in dados.columns:
            dados[coluna] = ""

    dados = dados[COLUNAS_PORTAL].copy()

    for coluna in COLUNAS_PORTAL:
        if coluna not in {"valor", "ativo"}:
            dados[coluna] = dados[coluna].fillna("").astype(str).str.strip()

    dados["valor"] = pd.to_numeric(dados["valor"], errors="coerce").fillna(0)
    dados["ativo"] = dados["ativo"].fillna(True)
    dados["ativo"] = dados["ativo"].apply(lambda valor: str(valor).strip().lower() not in {"false", "0", "nao", "não", "n"})
    return dados


def validar_dados_portal(df: pd.DataFrame) -> list[str]:
    """Retorna uma lista de problemas encontrados. Lista vazia significa dados validos."""
    erros: list[str] = []
    colunas_faltantes = [coluna for coluna in COLUNAS_OBRIGATORIAS if coluna not in df.columns]

    if colunas_faltantes:
        erros.append(f"Colunas faltantes: {', '.join(colunas_faltantes)}")
        return erros

    obrigatorias = ["empresa", "competencia", "tipo", "descricao", "valor", "status"]
    for coluna in obrigatorias:
        vazios = df[coluna].isna() | (df[coluna].astype(str).str.strip() == "")
        if vazios.any():
            erros.append(f"Coluna obrigatoria com valor vazio: {coluna}")

    tipos_invalidos = sorted(set(df["tipo"].astype(str).str.strip()) - TIPOS_VALIDOS - {""})
    if tipos_invalidos:
        erros.append(f"Tipos invalidos: {', '.join(tipos_invalidos)}")

    status_invalidos = sorted(set(df["status"].astype(str).str.strip().str.lower()) - STATUS_VALIDOS - {""})
    if status_invalidos:
        erros.append(f"Status invalidos: {', '.join(status_invalidos)}")

    valores_invalidos = pd.to_numeric(df["valor"], errors="coerce").isna()
    if valores_invalidos.any():
        erros.append("Existem valores nao numericos na coluna valor")

    for coluna_data in ["vencimento", "pagamento_recebimento"]:
        preenchidos = df[coluna_data].astype(str).str.strip() != ""
        datas_invalidas = pd.to_datetime(df.loc[preenchidos, coluna_data], errors="coerce").isna()
        if datas_invalidas.any():
            erros.append(f"Existem datas invalidas na coluna {coluna_data}")

    return erros


def salvar_dados_portal(df: pd.DataFrame, caminho_saida: str | Path = "data/dados_portal.csv") -> Path:
    """Normaliza, valida e salva o CSV que o app Streamlit le primeiro."""
    dados = normalizar_colunas_portal(df)
    erros = validar_dados_portal(dados)

    if erros:
        raise ValueError("Dados do portal invalidos:\n- " + "\n- ".join(erros))

    caminho = Path(caminho_saida)
    if not caminho.is_absolute():
        caminho = BASE_DIR / caminho
    caminho.parent.mkdir(parents=True, exist_ok=True)
    dados.to_csv(caminho, index=False, encoding="utf-8-sig")
    return caminho


def gerar_dados_demo_portal() -> pd.DataFrame:
    """Gera uma base pequena para testes do formato de exportacao."""
    registros = [
        {
            "empresa": "DMLIMA",
            "competencia": "2026-01",
            "tipo": "conta_a_pagar",
            "descricao": "Aluguel da sede",
            "fornecedor_cliente": "Imobiliaria Central",
            "vencimento": "2026-01-10",
            "pagamento_recebimento": "",
            "valor": 3200.00,
            "status": "aberto",
            "categoria": "Administrativo",
            "observacao": "Registro demonstrativo para testes",
            "documento": "AP-DML-001",
            "tipo_conta_codigo": "2.1.6.02.001",
            "tipo_conta_nome": "HONORARIOS CONTABEIS",
            "anexo_nome": "",
            "anexo_caminho": "",
            "codigo_pagamento": "",
            "criado_em": "",
            "criado_por": "",
            "excluido_em": "",
            "excluido_por": "",
            "ativo": True,
        },
        {
            "empresa": "DMLIMA",
            "competencia": "2026-01",
            "tipo": "receita",
            "descricao": "Servicos de consultoria",
            "fornecedor_cliente": "Cliente Alfa",
            "vencimento": "2026-01-15",
            "pagamento_recebimento": "2026-01-15",
            "valor": 12500.00,
            "status": "recebido",
            "categoria": "Receita operacional",
            "observacao": "Registro demonstrativo para testes",
            "documento": "NF-DML-001",
            "tipo_conta_codigo": "",
            "tipo_conta_nome": "",
            "anexo_nome": "",
            "anexo_caminho": "",
            "codigo_pagamento": "",
            "criado_em": "",
            "criado_por": "",
            "excluido_em": "",
            "excluido_por": "",
            "ativo": True,
        },
        {
            "empresa": "DMLIMA",
            "competencia": "2026-01",
            "tipo": "relatorio",
            "descricao": "Balancete mensal",
            "fornecedor_cliente": "Escritorio Contabil",
            "vencimento": "2026-01-31",
            "pagamento_recebimento": "",
            "valor": 0.00,
            "status": "disponível",
            "categoria": "Relatorios contabeis",
            "observacao": "Documento demonstrativo disponivel",
            "documento": "REL-DML-001",
            "tipo_conta_codigo": "",
            "tipo_conta_nome": "",
            "anexo_nome": "",
            "anexo_caminho": "",
            "codigo_pagamento": "",
            "criado_em": "",
            "criado_por": "",
            "excluido_em": "",
            "excluido_por": "",
            "ativo": True,
        },
    ]
    return normalizar_colunas_portal(pd.DataFrame(registros))


def exemplo_de_uso() -> Path:
    """Exemplo manual: gera data/dados_portal.csv a partir de dados demo."""
    df = gerar_dados_demo_portal()
    return salvar_dados_portal(df)


if __name__ == "__main__":
    saida = exemplo_de_uso()
    print(f"Arquivo gerado: {saida}")
