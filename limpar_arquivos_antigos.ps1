param(
  [int]$DiasRetencao = 30,
  [switch]$Executar
)

$pastas = @(
  @{ Nome = "entrada";    Caminho = ".\entrada";    Padrao = "*.xlsx" },
  @{ Nome = "historico";  Caminho = ".\historico";  Padrao = "*.xlsx" },
  @{ Nome = "alertas";    Caminho = ".\alertas";    Padrao = "*.json" },
  @{ Nome = "logs";       Caminho = ".\logs";       Padrao = "*.log" },
  @{ Nome = ".run_guard"; Caminho = ".\.run_guard"; Padrao = "*.last" }
)

$limite = (Get-Date).AddDays(-$DiasRetencao)
$totalArquivos = 0
$totalBytes = 0

Write-Host "Limpando arquivos com mais de $DiasRetencao dias (anteriores a $($limite.ToString('dd/MM/yyyy')))..."
if (-not $Executar) {
  Write-Host "Modo simulacao (nada sera apagado). Use -Executar para apagar de verdade."
}
Write-Host ""

foreach ($pasta in $pastas) {
  if (-not (Test-Path $pasta.Caminho)) {
    Write-Host "$($pasta.Nome): pasta nao encontrada, pulando."
    continue
  }

  $antigos = Get-ChildItem -Path $pasta.Caminho -Filter $pasta.Padrao -File | Where-Object { $_.LastWriteTime -lt $limite }

  if ($antigos.Count -eq 0) {
    Write-Host "$($pasta.Nome): nada mais velho que $DiasRetencao dias."
    continue
  }

  $bytesPasta = ($antigos | Measure-Object -Property Length -Sum).Sum
  $totalArquivos += $antigos.Count
  $totalBytes += $bytesPasta

  Write-Host "$($pasta.Nome): $($antigos.Count) arquivo(s), $([math]::Round($bytesPasta / 1MB, 1)) MB"

  if ($Executar) {
    $antigos | Remove-Item -Force
    Write-Host "  -> apagados."
  }
}

Write-Host ""
if ($Executar) {
  Write-Host "Total apagado: $totalArquivos arquivo(s), $([math]::Round($totalBytes / 1MB, 1)) MB"
}
else {
  Write-Host "Total que SERIA apagado: $totalArquivos arquivo(s), $([math]::Round($totalBytes / 1MB, 1)) MB"
  Write-Host "Rode novamente com -Executar para apagar de verdade."
}
