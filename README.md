# Portal Contábil do Cliente

Aplicativo web em Streamlit para um portal contábil simples, separado do sistema principal `dominio_dmls_08.py`.

Nesta versão, a função principal é **Contas a Pagar**:

- login por usuário e senha;
- seleção de competência após o login;
- lista de contas em aberto;
- inclusão manual de contas a pagar;
- upload de anexos, como boletos, notas fiscais e guias;
- exclusão lógica de contas, preservando auditoria;
- registro de data/hora e usuário responsável pela inclusão ou exclusão;
- tipos de conta baseados nos passivos monitorados no `dominio_dmls_08.py`.

O app não usa banco de dados. As alterações manuais são gravadas em `data/dados_portal.csv`, arquivo ignorado pelo Git para evitar envio acidental de dados operacionais.

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

No Windows, também há dois atalhos na pasta do projeto:

- `Iniciar Portal Local.bat`: inicia o Streamlit localmente.
- `Abrir Portal Local.url`: abre `http://localhost:8501` no navegador.

## Acessos De Teste

Use estes acessos demonstrativos para testar:

```text
DMLIMA / 123456
VICTOR / 123456
MHLOG  / 123456
MH BRASIL / 123456
```

## Publicação No Streamlit Community Cloud

1. Crie ou acesse o repositório no GitHub.
2. Envie os arquivos desta pasta para o repositório.
3. Acesse [Streamlit Community Cloud](https://streamlit.io/cloud).
4. Clique em **Create app** ou **New app**.
5. Selecione o repositório, a branch `main` e o arquivo principal `app.py`.
6. Clique em **Deploy**.

Se o app estiver dentro de uma subpasta do repositório, informe o caminho completo, por exemplo:

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

## Arquivos Operacionais Locais

Estes arquivos podem ser criados em uso local, mas não devem ser enviados ao GitHub:

- `data/dados_portal.csv`
- `data/anexos/`
- `.streamlit/secrets.toml`

## Segurança

Não suba dados reais sensíveis em repositório público.

Não coloque senhas reais no GitHub. Para produção, use `st.secrets` no Streamlit Community Cloud ou outro mecanismo seguro de autenticação.

Os anexos salvos localmente no Streamlit Community Cloud podem ser temporários, porque a hospedagem gratuita não é um armazenamento permanente. Para uso real com clientes, o próximo passo é guardar anexos em serviço externo seguro, como Google Drive controlado, S3 ou banco com storage.
