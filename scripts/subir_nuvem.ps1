$ErrorActionPreference = "Stop"

$Projeto = Split-Path -Parent $PSScriptRoot
Set-Location $Projeto

Write-Host ""
Write-Host "=== Portal Cliente Streamlit - Subir para nuvem ===" -ForegroundColor Green
Write-Host ""

$Git = Get-Command git -ErrorAction SilentlyContinue
if (-not $Git) {
    $GitDesktop = Get-ChildItem "$env:LOCALAPPDATA\GitHubDesktop\app-*\resources\app\git\cmd\git.exe" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($GitDesktop) {
        $env:PATH = "$(Split-Path -Parent $GitDesktop.FullName);$env:PATH"
        $Git = Get-Command git -ErrorAction SilentlyContinue
    }
}

Write-Host "1) Validando app.py..."
python -m py_compile app.py

Write-Host "2) Gerando pacote limpo de deploy..."
$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Dist = Join-Path $Projeto "deploy_streamlit_$Stamp"
$Zip = "$Dist.zip"

New-Item -ItemType Directory -Path $Dist -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $Dist ".streamlit") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $Dist "assets") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $Dist "data") -Force | Out-Null

Copy-Item -LiteralPath "app.py", "requirements.txt", "README.md", "config_clientes.json", "exportador_portal.py", ".gitignore" -Destination $Dist -Force
Copy-Item -LiteralPath ".streamlit\secrets.toml.example" -Destination (Join-Path $Dist ".streamlit") -Force
Copy-Item -LiteralPath "assets\mh_log_logo_app_header.png", "assets\mh_log_logo_app_512.png" -Destination (Join-Path $Dist "assets") -Force
Copy-Item -LiteralPath "data\dados_demo.csv" -Destination (Join-Path $Dist "data") -Force

Compress-Archive -Path (Join-Path $Dist "*") -DestinationPath $Zip -Force
Write-Host "Pacote criado:" -ForegroundColor Green
Write-Host $Zip
Write-Host ""

if (-not $Git) {
    Write-Host "Git nao esta instalado ou nao esta no PATH." -ForegroundColor Yellow
    Write-Host "Instale o Git, envie o projeto para o GitHub e depois publique no Streamlit Cloud."
    Start-Process "https://git-scm.com/download/win"
    Start-Process "https://share.streamlit.io/"
    Write-Host ""
    Read-Host "Pressione ENTER para fechar"
    exit 0
}

$Remote = ""
try {
    $Remote = git remote get-url origin 2>$null
} catch {
    $Remote = ""
}

if (-not $Remote) {
    Write-Host "Este repositorio ainda nao tem remote 'origin' configurado." -ForegroundColor Yellow
    Write-Host "Crie um repositorio no GitHub e rode:"
    Write-Host "git remote add origin URL_DO_REPOSITORIO"
    Start-Process "https://github.com/new"
    Start-Process "https://share.streamlit.io/"
    Write-Host ""
    Read-Host "Pressione ENTER para fechar"
    exit 0
}

Write-Host "3) Enviando arquivos para o GitHub..."
git add app.py requirements.txt README.md config_clientes.json exportador_portal.py .gitignore .streamlit/secrets.toml.example assets data/dados_demo.csv "Subir para Nuvem.bat" scripts/subir_nuvem.ps1

$Status = git status --porcelain
if ($Status) {
    $Mensagem = "Atualizacao portal cliente $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    git commit -m $Mensagem
} else {
    Write-Host "Nenhuma alteracao nova para commitar."
}

$Branch = git branch --show-current
if (-not $Branch) {
    $Branch = "main"
}

git push origin $Branch

Write-Host ""
Write-Host "Arquivos enviados para o GitHub com sucesso." -ForegroundColor Green
Write-Host "Abrindo Streamlit Cloud para publicar/atualizar o app..."
Start-Process "https://share.streamlit.io/"
Write-Host ""
Read-Host "Pressione ENTER para fechar"
