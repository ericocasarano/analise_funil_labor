param(
  [string]$Config = ".\automacao_config.json",
  [string]$JsonDir = "",
  [string]$Dax1Pattern = "dax1_funil_powerbi_*.json",
  [string]$Dax2Pattern = "dax2_itens_powerbi_*.json",
  [string]$Bloco3JsonDir = "",
  [string]$Bloco2Pattern = "bloco2_revisao_gestor_*.json",
  [string]$Bloco3Pattern = "bloco3_aprovacao_cliente_*.json",
  [string]$Oportunidades = "",
  [ValidateSet("sem_ruido", "com_ruido")]
  [string]$Visao = "com_ruido",
  [switch]$UsarBloco3,
  [switch]$ForceRun
)

$ErrorActionPreference = "Stop"

function ConvertTo-RunGuardName {
  param([string]$Value)
  return (($Value -replace '[^A-Za-z0-9_.-]', '_') -replace '_+', '_').Trim('_')
}

function Start-RunGuard {
  param(
    [string]$Name,
    [string]$Key,
    [int]$DuplicateWindowMinutes = 5,
    [switch]$Force
  )

  $guardDir = Join-Path $PSScriptRoot ".run_guard"
  if (-not (Test-Path $guardDir)) {
    New-Item -ItemType Directory -Path $guardDir | Out-Null
  }

  $safeName = ConvertTo-RunGuardName ("{0}_{1}" -f $Name, $Key)
  $script:RunGuardLockPath = Join-Path $guardDir ("{0}.lock" -f $safeName)
  $script:RunGuardMarkerPath = Join-Path $guardDir ("{0}.last" -f $safeName)

  if (-not $Force) {
    if (Test-Path $script:RunGuardLockPath) {
      $lockAge = (Get-Date) - (Get-Item $script:RunGuardLockPath).LastWriteTime
      if ($lockAge.TotalHours -lt 6) {
        Write-Host "Execucao identica ja esta em andamento. Bloqueando duplicidade."
        Write-Host "Use -ForceRun se realmente quiser ignorar a trava."
        return $false
      }
      Write-Host "Lock antigo encontrado. Removendo trava pendente antes de continuar."
      Remove-Item $script:RunGuardLockPath -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path $script:RunGuardMarkerPath) {
      $lastRun = Get-Item $script:RunGuardMarkerPath
      $age = (Get-Date) - $lastRun.LastWriteTime
      if ($age.TotalMinutes -lt $DuplicateWindowMinutes) {
        Write-Host ("Execucao identica concluida ha {0:N1} min. Bloqueando possivel duplicidade." -f $age.TotalMinutes)
        Write-Host "Use -ForceRun se realmente quiser rodar novamente."
        return $false
      }
    }
  }

  Set-Content -Path $script:RunGuardLockPath -Value ("Inicio={0}`nKey={1}" -f (Get-Date).ToString("s"), $Key) -Encoding UTF8
  return $true
}

function Complete-RunGuard {
  param([bool]$Succeeded)

  if ($script:RunGuardLockPath -and (Test-Path $script:RunGuardLockPath)) {
    Remove-Item $script:RunGuardLockPath -Force -ErrorAction SilentlyContinue
  }
  if ($Succeeded -and $script:RunGuardMarkerPath) {
    Set-Content -Path $script:RunGuardMarkerPath -Value ("Fim={0}" -f (Get-Date).ToString("s")) -Encoding UTF8
  }
}

function Get-LatestFile {
  param(
    [string]$Directory,
    [string]$Pattern
  )

  $item = Get-ChildItem $Directory -Filter $Pattern |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

  if (-not $item) {
    throw "Arquivo nao encontrado em $Directory com o padrao $Pattern"
  }

  return $item.FullName
}

function Get-LatestFileAfter {
  param(
    [string]$Directory,
    [string]$Pattern,
    [datetime]$After
  )

  if (-not (Test-Path $Directory)) {
    throw "Diretorio nao encontrado: $Directory"
  }

  $candidates = Get-ChildItem -LiteralPath $Directory -Filter $Pattern -File |
    Where-Object { $_.LastWriteTime -ge $After } |
    Sort-Object LastWriteTime -Descending

  if (-not $candidates) {
    throw "Nenhum arquivo encontrado em $Directory com o padrao $Pattern apos $After"
  }

  return $candidates[0].FullName
}

function Invoke-Python {
  param(
    [string]$PythonExecutable,
    [string[]]$Arguments
  )

  $pythonParts = $PythonExecutable -split '\s+'
  $pythonCmd = $pythonParts[0]
  $pythonCmdArgs = @()
  if ($pythonParts.Length -gt 1) {
    $pythonCmdArgs = $pythonParts[1..($pythonParts.Length - 1)]
  }

  & $pythonCmd @pythonCmdArgs @Arguments
  if ($LASTEXITCODE -ne 0) {
    throw "Comando Python falhou: $($Arguments -join ' ')"
  }
}

$inicioExecucao = Get-Date
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

$configObj = Get-Content $Config -Raw | ConvertFrom-Json
$pythonExe = $configObj.python_executable
if (-not $pythonExe) { $pythonExe = "py -3.12" }

if (-not $JsonDir) { $JsonDir = $configObj.paths.json_dir }
if (-not $JsonDir) {
  throw "Informe -JsonDir ou configure 'paths.json_dir' em $Config."
}
if ($UsarBloco3 -and -not $Bloco3JsonDir) { $Bloco3JsonDir = $configObj.paths.bloco3_json_dir }
if ($UsarBloco3 -and -not $Bloco3JsonDir) {
  throw "Informe -Bloco3JsonDir ou configure 'paths.bloco3_json_dir' em $Config."
}

Write-Host "Iniciando fluxo Painel Aprovacao do Gestor..."
Write-Host "Timestamp da rodada: $timestamp"
Write-Host "Pasta JSON:" $JsonDir
if ($UsarBloco3) {
  Write-Host "Pasta Bloco 3:" $Bloco3JsonDir
}
Write-Host "Visao do slide:" $Visao

$oportunidadesGuardPart = if ($Oportunidades) { $Oportunidades } else { "latest" }
$runGuardKey = "visao={0};oportunidades={1};usarbloco3={2}" -f $Visao, $oportunidadesGuardPart, $UsarBloco3.IsPresent
if (-not (Start-RunGuard -Name "rodar_fluxo_aprovacao_gestor" -Key $runGuardKey -Force:$ForceRun)) {
  return
}
$runGuardSucceeded = $false

$dax1SourceDir = $JsonDir
$dax1SourcePattern = $Dax1Pattern
if ($UsarBloco3) {
  $dax1SourceDir = $Bloco3JsonDir
  $dax1SourcePattern = $Bloco3Pattern
}

$dax1Json = Get-LatestFile -Directory $dax1SourceDir -Pattern $dax1SourcePattern
$dax2Json = Get-LatestFile -Directory $JsonDir -Pattern $Dax2Pattern
$bloco2Json = $null
if ($UsarBloco3) {
  $bloco2Json = Get-LatestFile -Directory $Bloco3JsonDir -Pattern $Bloco2Pattern
}

Write-Host "JSON DAX 1:" $dax1Json
Write-Host "JSON DAX 2:" $dax2Json
if ($UsarBloco3) {
  Write-Host "JSON Bloco 2:" $bloco2Json
}

$entradaDir = ".\entrada"
if (-not (Test-Path $entradaDir)) {
  New-Item -ItemType Directory -Path $entradaDir | Out-Null
}

$historicoDir = ".\historico"
if (-not (Test-Path $historicoDir)) {
  New-Item -ItemType Directory -Path $historicoDir | Out-Null
}

$alertasDir = ".\alertas"
if (-not (Test-Path $alertasDir)) {
  New-Item -ItemType Directory -Path $alertasDir | Out-Null
}

$arquivoRuido = Join-Path $entradaDir ("dax1_remocao_ruidos_aprovacao_gestor_{0}.xlsx" -f $timestamp)
$arquivoItens = Join-Path $entradaDir ("dax2_itens_orcamento_aprovacao_gestor_{0}.xlsx" -f $timestamp)
$arquivoKpiRevisao = Join-Path $entradaDir ("bloco2_revisao_gestor_aprovacao_gestor_{0}.xlsx" -f $timestamp)

Write-Host "Convertendo DAX 1..."
Invoke-Python -PythonExecutable $pythonExe -Arguments @(
  ".\tratar_dax1_json_power_automate.py",
  "-i", $dax1Json,
  "-o", $arquivoRuido
)

Write-Host "Convertendo DAX 2..."
Invoke-Python -PythonExecutable $pythonExe -Arguments @(
  ".\tratar_dax2_json_power_automate.py",
  "-i", $dax2Json,
  "-o", $arquivoItens
)

if ($UsarBloco3 -and $bloco2Json) {
  Write-Host "Convertendo Bloco 2..."
  Invoke-Python -PythonExecutable $pythonExe -Arguments @(
    ".\tratar_dax1_json_power_automate.py",
    "-i", $bloco2Json,
    "-o", $arquivoKpiRevisao
  )
}

$oportunidadesPath = $Oportunidades
if ($oportunidadesPath) {
  if (-not (Test-Path $oportunidadesPath)) {
    throw "Arquivo de oportunidades informado nao encontrado: $oportunidadesPath"
  }
}
elseif ($Visao -eq "sem_ruido") {
  Write-Host "Nenhum arquivo de oportunidades informado. Gerando oportunidades reais automaticamente..."

  $configObj.inputs.ruidos = $arquivoRuido
  $configObj.inputs.itens = $arquivoItens
  $configObj.filters.start = ""
  $configObj.filters.end = ""

  $configTemp = Join-Path $entradaDir ("automacao_config_aprovacao_gestor_{0}.json" -f $timestamp)
  $configObj | ConvertTo-Json -Depth 20 | Set-Content $configTemp -Encoding UTF8

  $runStep1Start = Get-Date
  Invoke-Python -PythonExecutable $pythonExe -Arguments @(
    ".\automacao_pipeline.py",
    "--config", $configTemp,
    "--resume-from", "step1",
    "--skip-step2"
  )

  $oportunidadesPath = Get-LatestFileAfter -Directory ".\historico" -Pattern "oportunidades_reais_auto_*.xlsx" -After $runStep1Start
}
else {
  Write-Host "Visao com_ruido selecionada. Nao sera gerado arquivo de oportunidades reais nesta rodada."
}

if ($oportunidadesPath) {
  Write-Host "Oportunidades reais:" $oportunidadesPath
}
Write-Host "Arquivo DAX 1 tratado:" $arquivoRuido
Write-Host "Arquivo DAX 2 tratado:" $arquivoItens
if ($UsarBloco3 -and $bloco2Json) {
  Write-Host "Arquivo KPI Revisao tratado:" $arquivoKpiRevisao
}

$outputExcel = Join-Path $historicoDir ("analise_aprovacao_gestor_{0}.xlsx" -f $timestamp)
$outputJson = Join-Path $alertasDir ("analise_aprovacao_gestor_{0}.json" -f $timestamp)

try {
  Write-Host "Gerando analise de aprovacao do gestor..."
  $analiseArgs = @(
    ".\gerar_analise_aprovacao_gestor.py",
    "--oportunidades-brutas", $arquivoRuido,
    "--itens", $arquivoItens,
    "--output", $outputExcel,
    "--output-json", $outputJson
  )
  if ($UsarBloco3 -and $bloco2Json) {
    $analiseArgs += @("--oportunidades-kpi", $arquivoKpiRevisao)
  }
  if ($oportunidadesPath) {
    $analiseArgs += @("--oportunidades", $oportunidadesPath)
  }
  Invoke-Python -PythonExecutable $pythonExe -Arguments $analiseArgs

  Write-Host "Atualizando slide interativo - Aprovacao Gestor..."
  Invoke-Python -PythonExecutable $pythonExe -Arguments @(
    ".\atualizar_slide_aprovacao_gestor_json.py",
    "-i", $outputJson,
    "--visao", $Visao
  )

  $runGuardSucceeded = $true
  $fimExecucao = Get-Date
  $duracao = New-TimeSpan -Start $inicioExecucao -End $fimExecucao

  Write-Host ""
  Write-Host "Fluxo Aprovacao Gestor concluido."
  Write-Host "Excel gerado:" $outputExcel
  Write-Host "JSON gerado:" $outputJson
  Write-Host "Slide atualizado: .\Slide de Win Rate Interativo\data\aprovacao_gestor_latest.js"
  Write-Host ("Tempo total: {0:mm}min {0:ss}s" -f $duracao)
}
finally {
  Complete-RunGuard -Succeeded:$runGuardSucceeded
}
