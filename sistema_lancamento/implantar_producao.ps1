# =============================================================================
# implantar_producao.ps1 - sobe o sistema de lancamento ambiental em producao
#
#   1) git clone https://github.com/ti-FG/portal_ambiental_qualidade-homologacao.git C:\Apps\portal_ambiental_qualidade
#   2) PowerShell como Administrador:
#        cd C:\Apps\portal_ambiental_qualidade\sistema_lancamento
#        .\implantar_producao.ps1
#      Na 1a vez ele cria o .env e abre o Bloco de Notas: preencha as linhas com ???,
#      salve e rode de novo.
#
# Para na primeira falha. Pode rodar de novo (atualizar o sistema): o banco ja
# aplicado e pulado e o servico e reiniciado.
#   -SemBanco    nao mexe no banco
#   -SemServico  nao cria/reinicia o servico (para testar)
#   -Sim         nao pergunta antes de aplicar a atualizacao no DW
# =============================================================================
param([switch]$SemBanco, [switch]$SemServico, [switch]$Sim)
$ErrorActionPreference = 'Stop'

$Raiz    = $PSScriptRoot
$Www     = Join-Path $Raiz 'www'
$EnvFile = Join-Path $Www '.env'
$Servico = 'FGAmbientalQualidade'

function Passo($t) { Write-Host ''; Write-Host "== $t" -ForegroundColor Cyan }
function Falha($t) { Write-Host "ERRO: $t" -ForegroundColor Red; exit 1 }

# ---------------------------------------------------------------- 1. ambiente
Passo '1/5 Conferindo o ambiente'
if ($Raiz -like '*\Dash_ambiental_qualidade\*') {
  Falha 'esta pasta fica dentro do site da 8080 (IIS serve tudo, o .env ficaria baixavel). Clone em C:\Apps\portal_ambiental_qualidade'
}
foreach ($c in 'node', 'npm') {
  if (-not (Get-Command $c -ErrorAction SilentlyContinue)) { Falha "$c nao encontrado" }
}
if (-not $SemServico) {
  $admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $admin) { Falha 'abra o PowerShell como Administrador' }
  if (-not (Get-Command nssm -ErrorAction SilentlyContinue)) { Falha 'nssm nao encontrado' }
}

if (-not (Test-Path $EnvFile)) {
  Copy-Item (Join-Path $Www '.env.exemplo') $EnvFile
  Write-Host "Criei $EnvFile" -ForegroundColor Yellow
  Write-Host 'Preencha as linhas com ??? (senhas), salve e rode este script de novo.' -ForegroundColor Yellow
  Start-Process notepad $EnvFile
  exit 0
}
$envTxt = Get-Content $EnvFile -Raw
$faltam = @($envTxt -split "`n" | Where-Object { $_ -notmatch '^\s*#' -and $_ -match '=\s*\?\?\?' } | ForEach-Object { ($_ -split '=')[0].Trim() })
if ($faltam.Count -gt 0) { Falha ("falta preencher no .env: " + ($faltam -join ', ') + "  ($EnvFile)") }
$porta = 8095
if ($envTxt -match '(?m)^\s*PORT\s*=\s*(\d+)') { $porta = [int]$Matches[1] }
Write-Host "ok. Porta do sistema: $porta"

# ------------------------------------------------------------ 2. dependencias
Passo '2/5 Instalando as dependencias (npm)'
Push-Location $Www
try {
  npm ci --omit=dev --no-audit --no-fund
  if ($LASTEXITCODE -ne 0) { Falha 'npm ci falhou' }
} finally { Pop-Location }

# ------------------------------------------------------------------- 3. banco
Passo '3/5 Banco (schema portal_ambiental_qualidade no DW)'
if ($SemBanco) {
  Write-Host 'pulado (-SemBanco)'
} else {
  $aplicar = Join-Path $Raiz 'db\aplicar_sql.js'
  node $aplicar (Join-Path $Raiz 'db\schema.sql')
  if ($LASTEXITCODE -ne 0) { Falha 'schema.sql falhou (confira host, usuario e senha do .env)' }

  $quer = $Sim
  if (-not $quer) {
    $resp = Read-Host 'Aplicar agora a atualizacao de dados no DW (uma unica vez, nao apaga nada)? [s/N]'
    $quer = ($resp -match '^(s|S|y|Y)')
  }
  if ($quer) {
    node $aplicar (Join-Path $Raiz 'db\atualizar_dw_para_homologacao.sql')
    if ($LASTEXITCODE -eq 3) { Write-Host 'a atualizacao ja tinha sido aplicada antes: tudo certo' }
    elseif ($LASTEXITCODE -ne 0) { Falha 'atualizar_dw_para_homologacao.sql falhou; nada foi gravado (a transacao foi desfeita)' }
  } else {
    Write-Host 'atualizacao de dados NAO aplicada (voce respondeu nao)' -ForegroundColor Yellow
  }
}

# ----------------------------------------------------------------- 4. servico
Passo '4/5 Servico do Windows'
if ($SemServico) {
  Write-Host 'pulado (-SemServico)'
} else {
  $node = (Get-Command node).Source
  $logs = Join-Path $Raiz 'logs'
  New-Item -ItemType Directory -Force $logs | Out-Null

  if (Get-Service $Servico -ErrorAction SilentlyContinue) {
    Stop-Service $Servico -ErrorAction SilentlyContinue
  } else {
    nssm install $Servico $node server.js | Out-Null
  }
  $dono = Get-NetTCPConnection -LocalPort $porta -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($dono) { Falha "a porta $porta ja esta em uso (PID $($dono.OwningProcess)); nada foi iniciado" }

  nssm set $Servico AppDirectory $Www | Out-Null
  nssm set $Servico DisplayName 'FG Lancamentos Ambientais' | Out-Null
  nssm set $Servico Start SERVICE_AUTO_START | Out-Null
  nssm set $Servico AppStdout (Join-Path $logs 'servidor.log') | Out-Null
  nssm set $Servico AppStderr (Join-Path $logs 'servidor.log') | Out-Null
  nssm set $Servico AppRotateFiles 1 | Out-Null
  nssm set $Servico AppRotateBytes 10485760 | Out-Null
  Start-Service $Servico
  Write-Host "servico $Servico iniciado"
}

# ----------------------------------------------------------------- 5. conferir
Passo '5/5 Conferindo'
if ($SemServico) {
  Write-Host 'pulado (-SemServico)'
} else {
  Start-Sleep -Seconds 10
  try {
    $r = Invoke-WebRequest "http://localhost:$porta/api/saude" -UseBasicParsing -TimeoutSec 20
    Write-Host $r.Content
    $j = $r.Content | ConvertFrom-Json
    if ($j.ok -and $j.banco -eq 'ok') { Write-Host 'saude: OK (sistema no ar e banco conectado)' -ForegroundColor Green }
    else { Falha 'o sistema subiu, mas o banco nao esta ok. Veja logs\servidor.log' }
    if (-not $j.ad) { Write-Host 'AVISO: login pelo AD nao configurado (LDAP_* do .env)' -ForegroundColor Yellow }
  } catch { Falha "nao respondeu em http://localhost:$porta/api/saude. Veja $logs\servidor.log" }

  foreach ($u in '.env', 'server.js') {
    $exposto = $true
    try { Invoke-WebRequest "http://localhost:$porta/$u" -UseBasicParsing -TimeoutSec 10 | Out-Null } catch { $exposto = $false }
    if ($exposto) { Falha "/$u esta ACESSIVEL pela rede. Pare o servico e avise a Natalia" }
  }
  Write-Host '/.env e /server.js nao sao acessiveis: OK' -ForegroundColor Green
}

Write-Host ''
Write-Host 'PRONTO.' -ForegroundColor Green
Write-Host "Falta: liberar a porta $porta no firewall, abrir http://10.1.1.151:$porta/ e entrar com um login do AD."
if ($envTxt -match '(?m)^\s*SYNC_AUTO\s*=\s*0') {
  Write-Host 'Depois de conferir, troque SYNC_AUTO=0 para 1 no .env e rode: nssm restart FGAmbientalQualidade'
}
