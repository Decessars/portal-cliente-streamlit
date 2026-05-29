# Portal Contabil do Cliente

Aplicativo web em Streamlit para um portal contabil simples, separado do sistema principal `dominio_dmls_08.py`.

Nesta versao, a funcao principal e **Contas a Pagar**:

- login por usuario e senha;
- dashboard por empresa apos o login;
- relogio em tempo real e clima local mediante permissao do navegador;
- lista de contas em aberto;
- download de contas em Excel formatado;
- inclusao manual de contas a pagar;
- upload de anexos, como boletos, notas fiscais e guias;
- exclusao logica de contas, preservando auditoria;
- registro de data/hora e usuario responsavel pela inclusao ou exclusao;
- tipos de conta baseados nos passivos monitorados no `dominio_dmls_08.py`.

## Persistencia Dos Dados

O app agora suporta Google Sheets como fonte principal de persistencia. Cada empresa pode ficar em uma aba propria da mesma planilha, por exemplo `MHLOG` e `MH_BRASIL`.

Configuracao esperada em `st.secrets`:

- `google_sheets.spreadsheet_id`
- `google_service_account` com a chave da service account

Quando essa configuracao existe, o app le e grava direto na planilha. Antes de cada gravacao, ele cria uma aba de backup no proprio Google Drive com prefixo `BACKUP__`.

Na primeira execucao com Google Sheets ativo, se a planilha estiver vazia e ainda existir um CSV local com dados, o app pode migrar essa base automaticamente para a aba da empresa.

Se a configuracao de Google Sheets nao existir, o app continua funcionando em modo local com CSV, mas isso nao e persistencia confiavel para Streamlit Cloud.

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

No Windows, tambem ha dois atalhos na pasta do projeto:

- `Iniciar Portal Local.bat`: inicia o Streamlit localmente.
- `Abrir Portal Local.url`: abre `http://localhost:8501` no navegador.

## Acessos De Teste

Use estes acessos demonstrativos para testar:

```text
DMLIMA / 123456
VITOR / 123456
ALEX / 123456
MH_FINANCEIRO / 123456
RAFAEL / 123456
JESSICA / 123456
```

Depois do login, o portal exibe o dashboard por empresa. `MHLOG` e `MH BRASIL` sao empresas/clientes, nao usuarios de acesso.

Perfis:

- `DMLIMA`: administrador, com acesso a tudo.
- `VITOR`: estagiario, pode operar o sistema.
- `ALEX`: dono da empresa, perfil de consulta.
- `MH_FINANCEIRO`: financeiro, pode visualizar e marcar contas como pagas.
- `RAFAEL` e `JESSICA`: socios da empresa, podem operar o sistema.

Perfis de consulta podem visualizar e baixar relatorios, mas nao podem incluir, duplicar, editar, pagar, importar ou excluir contas.

Quando um usuario entra com a senha padrao `123456`, o portal exibe um aviso para trocar a senha ou confirmar que deseja permanecer com ela. As senhas continuam registradas no arquivo de configuracao do portal.
Depois do login, tambem e possivel trocar a senha pelo botao `Trocar senha` na barra lateral.
O arquivo `usuarios_perfis.txt` fica na pasta do projeto apenas para consulta e e atualizado automaticamente a partir do `config_clientes.json`.

## Como Ativar No Drive

1. Crie uma planilha privada no Google Drive.
2. Copie o `spreadsheet_id` para `st.secrets`.
3. Crie uma `service account` no Google Cloud e compartilhe a planilha com o e-mail dela.
4. Preencha `google_service_account` em `st.secrets`.
5. Crie uma aba para cada empresa, ou deixe o app criar automaticamente.

O caminho `H:\\Meu Drive\\...` do seu computador e apenas a pasta local sincronizada pelo Google Drive. Ele nao substitui a persistencia real da nuvem no Streamlit Cloud, que depende da API do Google Sheets.

## Recuperacao De Dados

Se a aba operacional sumir ou voltar vazia, use o alerta da tela de Contas a Pagar para restaurar o backup mais recente da propria planilha.

Fluxo de recuperacao:

1. Abra a planilha configurada em `google_sheets.spreadsheet_id`.
2. Localize a aba de backup mais recente, com prefixo `BACKUP__`.
3. Restaure a aba da empresa a partir desse backup.
4. Reabra o portal.

Para uso local sem Google Sheets, o app ainda usa `data/empresas/<EMPRESA>/dados_portal.csv` e `data/backups/<EMPRESA>/`.

## Publicacao No Streamlit Community Cloud

1. Crie ou acesse o repositorio no GitHub.
2. Envie os arquivos desta pasta para o repositorio.
3. Acesse [Streamlit Community Cloud](https://streamlit.io/cloud).
4. Clique em **Create app** ou **New app**.
5. Selecione o repositorio, a branch `main` e o arquivo principal `app.py`.
6. Clique em **Deploy**.

Se o app estiver dentro de uma subpasta do repositorio, informe o caminho completo, por exemplo:

```text
portal_cliente_streamlit/app.py
```

## Estrutura Dos Arquivos

```text
portal_cliente_streamlit/
|-- app.py
|-- requirements.txt
|-- README.md
|-- .gitignore
|-- config_clientes.json
|-- exportador_portal.py
|-- .streamlit/
|   `-- secrets.toml.example
|-- assets/
|   |-- mh_log_logo_app_header.png
|   `-- mh_log_logo_app_512.png
`-- data/
    `-- dados_demo.csv
```

## Arquivos Operacionais Locais

Estes arquivos podem ser criados em uso local, mas nao devem ser enviados ao GitHub:

- `data/empresas/`
- `data/backups/`
- `data/anexos/`
- `.streamlit/secrets.toml`

## Seguranca

Nao suba dados reais sensiveis em repositorio publico.

Nao coloque senhas reais no GitHub. Para producao, use `st.secrets` no Streamlit Community Cloud ou outro mecanismo seguro de autenticacao.

Os anexos salvos localmente no Streamlit Community Cloud podem ser temporarios, porque a hospedagem gratuita nao e um armazenamento permanente. Para uso real com clientes, o proximo passo e guardar anexos em servico externo seguro, como Google Drive controlado, S3 ou banco com storage.

Para dados operacionais, a recomendacao futura continua sendo um banco de dados ou storage externo com retencao e backup. O modo Google Sheets cobre persistencia simples e compartilhada, mas nao substitui um banco relacional.
