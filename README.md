# Analise Funil

## Visao Geral
Esta automacao hoje parte de uma unica extracao do Power Automate e gera 2 saidas da mesma rodada:

- `Insight Funil`: resumo executivo principal do funil
- `Comparativo`: comparativo entre periodos com a mesma base da rodada
- `Resumo Mensal Funil`: consolidado mensal auxiliar para analise historica em modo comercial ou calendario
- `Painel Aprovacao do Gestor`: painel visual e analise executiva da jornada de revisao do gestor

Regra de ordem da publicacao:

1. primeiro publicar o `Insight Funil`
2. depois publicar o `Comparativo`

Modo de uso atual:

- o Power Automate executa as consultas DAX e grava os JSONs na pasta de extracao
- a execucao local acontece manualmente, por comando, em `Teste` ou `Producao`

Observacao sobre os nomes dos scripts:

- o fluxo principal local agora esta em `C:\analise_funil\rodar_fluxo_funil.ps1`
- ele concentra a execucao principal da rodada antes da etapa comparativa
- o painel de aprovacao do gestor roda por `C:\analise_funil\rodar_fluxo_aprovacao_gestor.ps1`

## 1. Fluxo Principal da Rodada
Objetivo: gerar o JSON final do resumo executivo do funil e enviar para a pasta de teste ou producao, de onde o Power Automate publica o Adaptive Card no Teams.

Etapas:

1. O Power Automate executa as consultas DAX no Power BI.
2. Os arquivos brutos sao salvos na pasta SharePoint `/Entrada/FunilExtracao`.
3. O PowerShell local le os arquivos mais recentes.
4. Os dados sao tratados e convertidos.
5. O pipeline tecnico gera a base consolidada da rodada.
6. O `Insight Funil` e gerado primeiro.
7. O JSON do `Insight Funil` e copiado para a pasta final de `Teste` ou `Producao`.
8. Um fluxo do Power Automate monitora essa pasta final e envia o card ao Teams.

Comportamento padrao do periodo no `Insight Funil`:

- se `-StartDate` e `-EndDate` nao forem informados, o fluxo principal usa:
  - inicio = primeiro dia do mes comercial atual
  - fim = ontem
- o JSON final do insight separa:
  - `mes_comercial`
  - `periodo`
  - `atualizado_em`
- a extracao do Power Automate pode conter uma janela mais ampla, mas o card principal considera apenas o recorte efetivo da rodada

Regra padrao das nao convertidas:

- As analises e rankings de itens consideram, por padrao, orcamentos com `Faturou = 0`, `Aprovado pelo Cliente = 0` e status `Cancelado por Inatividade`, `Orçamento Cancelado` ou `Em confecção`.
- `Aprovado pelo Cliente` corresponde a coluna `ETAPA 3 FUNIL Aprovados pelo Cliente` da base de oportunidades.
- O recorte padrao de revisao e `filters.tipo_perda = "todas"`, ou seja, nao filtra `Passou por Revisao Gestor`.
- Para rodadas especificas, o script de itens ainda aceita `--tipo_perda todas`, `--tipo_perda revisadas` ou `--tipo_perda sem_revisao`.

Saida do step 2 (`itens_perdas_reais_auto_*.xlsx`):

- `Base_Nao_Convertidas`: orcamentos `Faturou = 0`, `Aprovado pelo Cliente = 0` e status `Cancelado por Inatividade`, `Orçamento Cancelado` ou `Em confecção`, respeitando o recorte de revisao informado.
- `Itens_Nao_Convertidas`: itens dos orcamentos da base nao convertida.
- `Base_Faturados` e `Itens_Faturados`: orcamentos com `Faturou = 1`, no mesmo recorte de revisao.
- Rankings finais por tipo de cliente: `Ranking_Geral_Final`, `Ranking_Geral_Revenda`, `Ranking_Geral_Final_Faturados` e `Ranking_Geral_Revenda_Faturados`.
- Analise de preco: `Analise_Preco_Itens`.

Abas removidas para simplificar a leitura do arquivo:

- `Ranking_Itens_Geral`
- `Ranking_Itens_Geral_Tipo`
- `Ranking_NC_Recusa_Tipo`
- `Ranking_NC_Recusa_Final`
- `Ranking_NC_Recusa_Revenda`
- `Ranking_Itens_Tipo_Faturados`
- `Base_NC_Recusa_Cliente`
- `Itens_NC_Recusa_Cliente`
- `Analise_Preco_Recusa`

Colunas de preco nos rankings:

- `Preco_Ponderado`: valor total dividido pelo volume total do item.
- `Preco_Media_Orcamento`: media dos precos calculados por orcamento.
- `Preco_Mediana_Orcamento`: mediana dos precos calculados por orcamento.

Na analise de preco, a `Classificacao_Preco` usa a diferenca pela mediana (`Dif_Preco_Mediana_%`) com limite de 3%. A media ponderada permanece disponivel como referencia financeira/volume.

Arquivos esperados na extracao:

- `dax1_funil_powerbi_yyyyMMdd_HHmmss.json`
- `dax2_itens_powerbi_yyyyMMdd_HHmmss.json`

Observacao sobre win rate:

- o win rate atual do funil continua baseado na `Data de Criacao`
- o pipeline tambem passa a calcular um win rate adicional por `Data de Faturamento`
- nesse segundo recorte, a data considerada por oportunidade e:
  - `Data de Faturamento`, quando existir
  - `Data de Criacao`, quando a data de faturamento estiver em branco
- para auditoria desse segundo recorte, o Excel passa a gerar a aba `Lista_WR_Data_Fat`
- `Comparativo_Geral_Total` permanece com a leitura original por data de criacao
- a leitura adicional fica separada na aba `Comp_Geral_Total_Data_Fat`

Script principal atual:

- `C:\analise_funil\rodar_fluxo_funil.ps1`

Comando de teste:

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_funil.ps1 -Ambiente Teste
```

Comando de producao:

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_funil.ps1 -Ambiente Producao
```

Comando sem envio ao SharePoint:

```powershell
.\rodar_fluxo_funil.ps1 -Ambiente Teste -NoSharePoint
```

## 2. Fluxo Comparativo
Objetivo: gerar o JSON comparativo entre 2 periodos e enviar para a pasta final de teste ou producao, de onde o Power Automate publica o card comparativo no Teams.

Etapas:

1. O comparativo usa a mesma base da rodada utilizada pelo `Insight Funil`.
2. O periodo atual e o periodo anterior proporcional sao recortados a partir dessa base.
3. Sao gerados 2 resumos independentes.
4. O script consolida os 2 resumos em um JSON comparativo.
5. O JSON comparativo e copiado para a pasta final de `Teste` ou `Producao`.
6. Um fluxo do Power Automate monitora essa pasta final e envia o card comparativo ao Teams.

Comportamento padrao do periodo no `Comparativo`:

- se `-StartDate` e `-EndDate` nao forem informados, o `Periodo Atual` usa:
  - inicio = primeiro dia do mes comercial atual
  - fim = ontem
- o `Periodo Anterior` usa:
  - inicio = primeiro dia do mes comercial anterior
  - fim = mesma quantidade de dias uteis observados no periodo atual, dentro do mes comercial anterior
- o JSON comparativo separa:
  - `periodo_a.label`
  - `periodo_a.dias_uteis`
  - `periodo_b.label`
  - `periodo_b.dias_uteis`
  - `atualizado_em`

Observacao sobre nao convertidas no comparativo:

- Quando o comparativo executa o pipeline auxiliar, ele tambem respeita a regra padrao de nao convertidas: `Faturou = 0`, `Aprovado pelo Cliente = 0`, status de recusa/cancelamento e `tipo_perda = todas`.
- A troca entre `Teste` e `Producao` muda apenas o destino de publicacao; a regra de calculo continua a mesma.

Script principal:

- `C:\analise_funil\rodar_fluxo_comparativo_funil.ps1`

Comando de teste:

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste
```

Comando de producao:

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Producao
```

## 3. Fluxos Power Automate
Recomendacao de organizacao:

- `Fluxo 1 - Extracao DAX Power BI`
  Funcao: executar DAX 1 e DAX 2 e salvar arquivos brutos em `/Entrada/FunilExtracao`.
  Observacao: a extracao deve cobrir a janela necessaria para atender o `Insight Funil` e o `Comparativo` na mesma rodada.

- `Fluxo 2 - Publicacao Insight Diario`
  Funcao: monitorar a pasta final do `Insight Funil` e postar o Adaptive Card no Teams.

- `Fluxo 3 - Publicacao Comparativo`
  Funcao: monitorar a pasta final de comparativo e postar o Adaptive Card comparativo no Teams.

## 4. Resumo Mensal Auxiliar
Objetivo: gerar uma planilha mensal historica com os mesmos conceitos de remocao de ruido usados no funil principal, sem depender do Excel consolidado da rodada.

Script principal:

- `C:\analise_funil\gerar_resumo_mensal_funil.py`

Compatibilidade:

- `C:\analise_funil\gerar_resumo_mensal_comercial.py`
  - nome antigo mantido como atalho para o script principal atual

Entradas aceitas:

- `dax1_funil_powerbi_*.json` e `dax2_itens_powerbi_*.json` diretamente da pasta de extracao
- arquivos tratados em `xlsx/xls/csv`

Saida:

- `historico\resumo_mensal_funil_YYYYMMDD_HHMMSS.xlsx`
  - aba `Resumo_Mensal_Funil`

Colunas geradas:

- `Mês/Ano Comercial`
- `Periodo`
- `Enviados_Qtd (Sem Ruído)`
- `Enviados_Valor (Sem Ruído)`
- `Faturado_Qtd (Sem Ruído)`
- `Faturado_Valor (Sem Ruído)`
- `Nao_Convertidas_Qtd (Sem Ruído)`
- `Nao_Convertidas_Valor (Sem Ruído)`
- `Win Rate (Volume) % (Sem Ruído)`
- `Win Rate (Valor) % (Sem Ruído)`

Modos disponiveis:

- `comercial`
  - modo padrao
  - cada linha representa um mes comercial
- `calendario`
  - ativado com `--modo calendario`
  - cada linha representa um mes calendario civil

Observacao sobre meses parciais:

- quando a base informada nao cobre o ciclo inteiro, o script mantem a linha e informa no console quais meses ficaram parciais
- isso costuma acontecer no primeiro mes da janela e no ultimo mes em andamento

Exemplo de mes calendario:

```powershell
cd C:\analise_funil
py -3.12 .\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --modo calendario --start 2026-04-01 --end 2026-04-30 -o resumo_calendario_abril
```

Exemplo de mes comercial:

```powershell
cd C:\analise_funil
py -3.12 .\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --start 2026-03-31 --end 2026-04-29 -o resumo_comercial_abril
```

Exemplo usando os JSONs mais recentes da pasta de extracao:

```powershell
cd C:\analise_funil
$jsonDir = "C:\CAMINHO\PARA\Entrada\FunilExtracao"
$dax1 = Get-ChildItem -LiteralPath $jsonDir -Filter "dax1_funil_powerbi_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$dax2 = Get-ChildItem -LiteralPath $jsonDir -Filter "dax2_itens_powerbi_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
py -3.12 .\gerar_resumo_mensal_funil.py -i $dax1.FullName -it $dax2.FullName --modo calendario --start 2025-12-01 --end 2026-04-30 -o resumo_calendario_dez_abr
```

## 5. Estrutura de Pastas
Sugestao de organizacao:

- Entrada bruta:
  `/Entrada/FunilExtracao`

- Saida insight teste:
  pasta configurada em `insight_json_dir` do ambiente `Teste`

- Saida insight producao:
  pasta configurada em `insight_json_dir` do ambiente `Producao`

- Saida comparativo teste:
  pasta configurada em `comparativo_json_dir` do ambiente `Teste`

- Saida comparativo producao:
  pasta configurada em `comparativo_json_dir` do ambiente `Producao`

## 6. Regra de Uso
Use assim:

- `rodar_fluxo_funil.ps1`
  Quando quiser executar a rodada principal do funil e gerar o `Insight Funil`

- `rodar_fluxo_comparativo_funil.ps1`
  Quando quiser gerar o comparativo depois da rodada principal

- `gerar_resumo_mensal_funil.py`
  Quando quiser gerar uma visao historica mensal auxiliar da base, em mes comercial ou calendario

Observacao operacional:

- hoje o uso previsto e manual
- primeiro rode o `Insight Funil`
- depois rode o `Comparativo`
- a extracao base ja deve ter sido gerada pelo Power Automate antes dos comandos locais
- no modo padrao:
  - o `Insight Funil` considera mes comercial atual ate ontem
  - o `Comparativo` considera mes comercial atual ate ontem versus periodo equivalente do mes comercial anterior por dias uteis
  - o `Resumo Mensal Funil` usa o modo `comercial`, salvo quando `--modo calendario` for informado

Observacao sobre leitura temporal nos cards:

- no `Insight Funil`, o card mostra:
  - `Mes Comercial`
  - `Periodo analisado`
  - `Atualizado em`
- no `Comparativo`, o card mostra:
  - `Periodo Anterior`
  - `Periodo Atual`
  - `Atualizado em`

## 7. Resumo Operacional
Insight Funil em teste:

```powershell
.\rodar_fluxo_funil.ps1 -Ambiente Teste
```

Insight Funil em producao:

```powershell
.\rodar_fluxo_funil.ps1 -Ambiente Producao
```

Comparativo teste:

```powershell
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste
```

Comparativo producao:

```powershell
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Producao
```

Resumo mensal em mes calendario:

```powershell
.\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --modo calendario --start 2026-04-01 --end 2026-04-30 -o resumo_calendario_abril
```

Resumo mensal em mes comercial:

```powershell
.\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --start 2026-03-31 --end 2026-04-29 -o resumo_comercial_abril
```
