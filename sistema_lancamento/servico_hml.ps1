# =============================================================================
# servico_hml.ps1 — Lançamentos Ambientais em HOMOLOGAÇÃO como serviço (NSSM)
#
#   pwsh -ExecutionPolicy Bypass -File servico_hml.ps1      (como administrador)
#
# Mesmo modelo do Portfólio de Projetos (painel_project_manager\www\servico_hml.ps1)
# e dos memdb do Quality/Jurídico: dois serviços, início automático, LocalSystem,
# log com rotação de 10 MB.
#   FGAmbientalQualidadeMemdb  hml\memdb.mjs   banco PGlite em 127.0.0.1:15439
#                                              (persistido em hml\dados)
#   FGAmbientalQualidadeHML    www\server.js   sistema em http://10.1.1.155:8095
#                                              (depende do memdb)
# Rodar de novo é seguro: atualiza a configuração e reinicia.
# Mudança no servidor = "nssm restart FGAmbientalQualidadeHML" (as sessões
# ficam no banco: ninguém é deslogado).
# Voltar o banco ao estado da planilha: parar os dois, apagar hml\dados, subir.
# =============================================================================
$ErrorActionPreference = 'Stop'
$Raiz = $PSScriptRoot
$Node = (Get-Command node).Source
$Logs = Join-Path $Raiz 'logs'
New-Item -ItemType Directory -Force $Logs | Out-Null

function Instalar($nome, $exibir, $descricao, $dir, $script, $log, $depende) {
  $svc = Get-Service $nome -ErrorAction SilentlyContinue
  if (-not $svc) { Write-Host "instalando $nome"; nssm install $nome $Node $script | Out-Null }
  else { Write-Host "$nome já existe: atualizando a configuração" }
  nssm set $nome AppDirectory $dir | Out-Null
  nssm set $nome DisplayName $exibir | Out-Null
  nssm set $nome Description $descricao | Out-Null
  nssm set $nome Start SERVICE_AUTO_START | Out-Null
  nssm set $nome AppStdout $log | Out-Null
  nssm set $nome AppStderr $log | Out-Null
  nssm set $nome AppStdoutCreationDisposition 4 | Out-Null
  nssm set $nome AppStderrCreationDisposition 4 | Out-Null
  nssm set $nome AppRotateFiles 1 | Out-Null
  nssm set $nome AppRotateBytes 10485760 | Out-Null
  # PGlite fecha o banco com SIGINT/Ctrl+C: dar tempo para gravar em disco
  nssm set $nome AppStopMethodConsole 15000 | Out-Null
  if ($depende) { nssm set $nome DependOnService $depende | Out-Null }
}

foreach ($p in 8095, 15439) {
  $dono = (Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue).OwningProcess | Select-Object -First 1
  if ($dono) {
    $cmd = (Get-CimInstance Win32_Process -Filter "ProcessId=$dono").CommandLine
    $ehServico = (Get-CimInstance Win32_Service | Where-Object ProcessId -eq (Get-CimInstance Win32_Process -Filter "ProcessId=$dono").ParentProcessId)
    if (-not $ehServico) {
      if ($cmd -notmatch 'server\.js|memdb\.mjs') { throw "a porta $p é de outro programa ($cmd); nada foi parado" }
      Write-Host "parando o node aberto à mão na porta $p (PID $dono)"
      Stop-Process -Id $dono -Force
      Start-Sleep -Seconds 2
    }
  }
}

Instalar 'FGAmbientalQualidadeMemdb' 'FG Lancamentos Ambientais - banco (HML)' `
  'Banco PGlite da homologacao dos Lancamentos Ambientais (127.0.0.1:15439, persistido em hml\dados)' `
  (Join-Path $Raiz 'hml') 'memdb.mjs' (Join-Path $Logs 'memdb.log') $null
Instalar 'FGAmbientalQualidadeHML' 'FG Lancamentos Ambientais (HML)' `
  'Lancamentos Ambientais da Qualidade - homologacao (porta 8095)' `
  (Join-Path $Raiz 'www') 'server.js' (Join-Path $Logs 'servidor.log') 'FGAmbientalQualidadeMemdb'

nssm restart FGAmbientalQualidadeMemdb | Out-Null
Start-Sleep -Seconds 8
nssm restart FGAmbientalQualidadeHML | Out-Null
Start-Sleep -Seconds 4

try {
  $r = Invoke-WebRequest -UseBasicParsing 'http://localhost:8095/api/saude' -TimeoutSec 20 -SkipHttpErrorCheck
  Write-Host "/api/saude: HTTP $($r.StatusCode) $($r.Content)"
} catch { Write-Host "/api/saude não respondeu: $($_.Exception.Message). Ver $Logs"; exit 1 }
