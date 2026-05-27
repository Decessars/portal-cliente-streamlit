# Portal Contábil do Cliente

Aplicativo web gratuito em Streamlit para demonstrar um portal contábil do cliente, com login simples, filtros por empresa e competência, resumo financeiro, contas a pagar, contas pagas, receitas, impostos e relatórios.

O projeto usa apenas arquivos locais demonstrativos (`CSV` ou `XLSX`) e não usa banco de dados. Se o arquivo de dados não existir, o app cria automaticamente uma base demo em `data/dados_demo.csv`.

Para uma integração futura, o app procura primeiro `data/dados_portal.csv`. Esse arquivo pode ser gerado manualmente ou pelo módulo `exportador_portal.py`. Se ele não existir, o app usa `data/dados_demo.csv`.

## Como Rodar Localmente

```bash
cd portal_cliente_streamlit
pip install -r requirements.txt
streamlit run app.py
```

Depois acesse:

```text
http://localhost:8501
```

## Publicação No Streamlit Community Cloud

1. Crie um repositório no GitHub.
2. Envie os arquivos desta pasta para o repositório.
3. Acesse [Streamlit Community Cloud](https://streamlit.io/cloud).
4. Clique em **New app**.
5. Selecione o repositório e a branch.
6. Em **Main file path**, informe `app.py`.
7. Clique em **Deploy**.

Se você publicar o projeto dentro de uma subpasta do repositório, informe o caminho completo, por exemplo:

```text
portal_cliente_streamlit/app.py
```

## Estrutura Dos Arquivos

```text
portal_cliente_streamlit/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── config_clientes.json
├── exportador_portal.py
├── .streamlit/
│   └── secrets.toml.example
├── assets/
│   ├── mh_log_logo_app_header.png
│   └── mh_log_logo_app_512.png
└── data/
    └── dados_demo.csv
```

## Integração Futura

O arquivo `exportador_portal.py` prepara a padronização dos dados que poderão vir de outro sistema no futuro. Ele não altera nem importa o sistema principal.

Funções disponíveis:

- `normalizar_colunas_portal(df)`
- `validar_dados_portal(df)`
- `salvar_dados_portal(df, caminho_saida="data/dados_portal.csv")`
- `gerar_dados_demo_portal()`
- `exemplo_de_uso()`

## Dados Demonstrativos

O arquivo `data/dados_demo.csv` usa as colunas:

- empresa
- competencia
- tipo
- descricao
- fornecedor_cliente
- vencimento
- pagamento_recebimento
- valor
- status
- categoria
- observacao
- documento

## Segurança

Não suba dados reais sensíveis em repositório público.

Não coloque senhas reais no GitHub. Use apenas senhas demonstrativas para testes locais. Em uma versão de produção, mova credenciais para `st.secrets` no Streamlit Community Cloud ou para outro mecanismo seguro de autenticação.
