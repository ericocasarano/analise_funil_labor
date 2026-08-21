param(
  [ValidateSet("Teste", "Producao")]
  [string]$Ambiente = "Producao"
)

Set-Location -Path $PSScriptRoot

$logDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logDir)) {
  New-Item -ItemType Directory -Path $logDir | Out-Null
}
$logFile = Join-Path $logDir ("agendamento_fluxos_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
Start-Transcript -Path $logFile -Append | Out-Null

try {
  Write-Host "=== Iniciando Fluxo 1 (Insight Funil) - $Ambiente ==="
  try {
    & ".\rodar_fluxo_funil.ps1" -Ambiente $Ambiente -SkipItensPerdas
  }
  catch {
    Write-Host "ERRO no Fluxo 1: $_"
    Write-Host "Fluxo 2 nao sera executado."
    throw
  }

  Write-Host ""
  Write-Host "=== Fluxo 1 concluido. Iniciando Fluxo 2 (Comparativo) - $Ambiente ==="
  try {
    & ".\rodar_fluxo_comparativo_funil.ps1" -Ambiente $Ambiente -SkipItensPerdas
  }
  catch {
    Write-Host "ERRO no Fluxo 2: $_"
    throw
  }

  Write-Host ""
  Write-Host "=== Fluxo 1 e Fluxo 2 concluidos com sucesso. ==="
}
finally {
  Stop-Transcript | Out-Null
}
