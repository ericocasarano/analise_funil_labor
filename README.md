# Analise Funil

## Visao Geral

Esta automacao parte de uma unica extracao do Power Automate e gera 2 saidas da mesma rodada:

- `Insight Funil`: resumo executivo principal do funil
- `Comparativo`: comparativo entre periodos com a mesma base da rodada
- `Resumo Mensal Funil`: consolidado mensal auxiliar para analise historica em modo comercial ou calendario

Regra de ordem da publicacao:

1. primeiro publicar o `Insight Funil`
2. depois publicar o `Comparativo`

No desenho atual da automacao local:

- `rodar_fluxo_funil.ps1` e o fluxo principal da rodada
- `rodar_fluxo_comparativo_funil.ps1` gera o comparativo

Modo de operacao atual:

- o Power Automate executa as consultas DAX no Power BI e grava os JSONs na pasta de extracao sincronizada
- a execucao local dos scripts acontece manualmente, por comando, em `Teste` ou `Producao`

### Estrategia adotada

A extracao base da rodada vem do Power Automate, com os arquivos JSON gravados na pasta de entrada sincronizada. Beneficios:

- separa a extracao do processamento local
- permite usar a mesma base para o `Insight Funil` e para o `Comparativo`
- reduz dependencia de login manual no Power BI para a operacao do dia a dia

## Ambientes

Os envios podem ser separados por ambiente:

### Producao

- Site SharePoint: `https://<seu-tenant>.sharepoint.com/sites/<seu-site>`
- Pasta resumo: `Documentos\Producao\FunilInsights`
- Pasta comparativo: `Documentos\Producao\FunilComparativos`

### Teste

- Site SharePoint: `https://<seu-tenant>.sharepoint.com/sites/<seu-site>`
- Pasta resumo: `Documentos\Teste\FunilInsights_teste`
- Pasta comparativo: `Documentos\Teste\FunilComparativos_teste`

## 1. Fluxo Principal da Rodada (Insight Funil)

Objetivo: gerar o JSON final do resumo executivo do funil e enviar para a pasta de teste ou producao, de onde o Power Automate publica o Adaptive Card no Teams.

### O que o fluxo faz

1. leitura dos arquivos JSON mais recentes gerados pelo Power Automate em `/Entrada/FunilExtracao`
2. tratamento dos arquivos em `entrada/`
3. execucao do pipeline Python
4. geracao do JSON do `Insight Funil`
5. copia do JSON para a pasta sincronizada do SharePoint / Teams
6. disparo do Power Automate para postar o Adaptive Card do insight
7. exibicao do tempo total da rodada

### Arquivos envolvidos

- `rodar_fluxo_funil.ps1`
- `tratar_dax_orcamentos_json_power_automate.py`
- `tratar_dax_itens_json_power_automate.py`
- `automacao_pipeline.py`
- `gerar_resumo_insight_json.py`

### Saidas principais

- `entrada\dax_orcamentos_tratado_power_automate_YYYYMMDD_HHMMSS.xlsx`
- `entrada\dax_itens_tratado_power_automate_YYYYMMDD_HHMMSS.xlsx`
- `historico\oportunidades_reais_auto_YYYYMMDD_HHMMSS.xlsx`
- `historico\itens_perdas_reais_auto_YYYYMMDD_HHMMSS.xlsx`
- `alertas\resumo_insight_card_teams_YYYYMMDD_HHMMSS.json`
- `.ultima_base_funil.json` (na raiz do projeto)
  - "bilhete" com o caminho exato do DAX1/DAX2 usados nessa rodada
  - lido pelo `rodar_fluxo_comparativo_funil.ps1` pra garantir que o Comparativo nasca da mesma base (ver Fluxo 2)
  - nao entra no Git (esta no `.gitignore`), e regenerado a cada rodada

### Regra padrao dos nao convertidos

Por padrao, o arquivo `historico\itens_perdas_reais_auto_YYYYMMDD_HHMMSS.xlsx` considera orcamentos com `Faturou = 0`, `Aprovado pelo Cliente = 0` e status de recusa/cancelamento, sem filtrar revisao do gestor.

No arquivo de oportunidades, `Aprovado pelo Cliente` vem da coluna `ETAPA 3 FUNIL Aprovados pelo Cliente`. Portanto, o padrao do step 2 remove da base de nao convertidas os casos que ja foram aprovados pelo cliente mas ainda nao faturaram.

Status considerados em `Base_Nao_Convertidas` e `Itens_Nao_Convertidas`:

- `Cancelado por Inatividade`
- `Orçamento Cancelado`
- `Em confecção`

A configuracao fica em `automacao_config.json`:

```json
"filters": {
  "tipo_perda": "todas"
}
```

Opcoes aceitas pelo script de itens nao convertidos: `todas`, `revisadas`, `sem_revisao`.

### Saida detalhada do step 2

O step 2 gera o arquivo `historico\itens_perdas_reais_auto_YYYYMMDD_HHMMSS.xlsx`.

Abas principais:

- `Base_Nao_Convertidas`: orcamentos com `Faturou = 0`, `Aprovado pelo Cliente = 0` e status de recusa/cancelamento, respeitando `tipo_perda`.
- `Itens_Nao_Convertidas`: itens desses orcamentos.
- `Base_Faturados`: orcamentos com `Faturou = 1`, no mesmo recorte de revisao.
- `Itens_Faturados`: itens dos orcamentos faturados.

Rankings mantidos: `Ranking_Geral_Final`, `Ranking_Geral_Revenda`, `Ranking_Geral_Final_Faturados`, `Ranking_Geral_Revenda_Faturados`, `Ranking_Itens_Por_Vendedor`, `Ranking_Itens_Por_Mes`, `Ranking_Itens_Vendedor_Mes`, abas `Top10_<Vendedor>`.

Abas removidas para simplificar o arquivo: `Ranking_Itens_Geral`, `Ranking_Itens_Geral_Tipo`, `Ranking_NC_Recusa_Tipo`, `Ranking_NC_Recusa_Final`, `Ranking_NC_Recusa_Revenda`, `Base_NC_Recusa_Cliente`, `Itens_NC_Recusa_Cliente`, `Analise_Preco_Recusa`, `Ranking_Itens_Tipo_Faturados`.

### Precos e classificacao no step 2

Colunas de preco nos rankings:

- `Preco_Ponderado`: `Valor_Total / Volume_Total`.
- `Preco_Media_Orcamento`: media dos precos calculados por orcamento.
- `Preco_Mediana_Orcamento`: mediana dos precos calculados por orcamento.

Na aba `Analise_Preco_Itens`, as mesmas visoes aparecem separadas para faturados e nao convertidos.

A `Classificacao_Preco` usa a mediana por orcamento como referencia principal (limite de 3%):

- `Dif_Preco_Mediana_% >= 3%` e `Win_Rate_Item_% < 50%`: `Possivel sensibilidade a preco`.
- `Dif_Preco_Mediana_% >= 3%`: `Atencao: preco nao convertido maior`.
- `Dif_Preco_Mediana_% <= -3%`: `Nao convertido com preco menor`.
- diferenca entre `-3%` e `+3%`: `Preco similar`.
- menos de 3 orcamentos faturados ou menos de 3 nao convertidos: `Amostra insuficiente`.

A media ponderada permanece disponivel como referencia financeira/volume.

### Pipeline tecnico atualmente utilizado

- `step1`: `gerar_oportunidades_reais_codes.py` — consolidacao das oportunidades reais
- `step2`: `gerar_itens_perdas_reais.py` — nao convertidos e rankings de itens

### Arquivos esperados na extracao

- `dax_query_base_orcamentos_funil_yyyyMMdd_HHmmss.json`
- `dax_query_base_itens_funil_yyyyMMdd_HHmmss.json`

### Observacao sobre win rate

- o win rate atual do funil continua baseado na `Data de Criacao`
- o pipeline tambem calcula um win rate adicional por `Data de Faturamento`
- nesse segundo recorte, a data considerada por oportunidade e:
  - `Data de Faturamento`, quando existir
  - `Data de Criacao`, quando a data de faturamento estiver em branco
- para auditoria desse segundo recorte, o Excel gera a aba `Lista_WR_Data_Fat`
- `Comparativo_Geral_Total` permanece com a leitura original por data de criacao
- a leitura adicional fica separada na aba `Comp_Geral_Total_Data_Fat`

### Regra de mes comercial no fluxo principal

- o mes comercial fecha no penultimo dia util do mes
- o dia seguinte ao fechamento passa a pertencer ao mes comercial seguinte
- o JSON do `Insight Funil` carrega `titulo`, `mes_comercial`, `periodo`, `atualizado_em`
- isso permite ao Adaptive Card diferenciar o ciclo comercial da rodada da parcela efetivamente analisada ate ontem

### Parametros disponiveis

- `-Ambiente Teste|Producao` — define a pasta sincronizada do SharePoint de destino
- `-StartDate YYYY-MM-DD` — define a data inicial da extracao
- `-EndDate YYYY-MM-DD` — define a data final da extracao
- `-NoSharePoint` — gera o JSON final apenas localmente, sem copiar para a pasta sincronizada

### Comportamento padrao

Se `-StartDate` e `-EndDate` nao forem informados:

- o script usa os arquivos JSON mais recentes disponiveis na pasta de extracao
- inicio = primeiro dia do mes comercial atual, fim = ontem
- o `Insight Funil` e gerado somente com esse recorte, mesmo que a extracao contenha uma janela mais ampla para atender o comparativo

### Comandos

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_funil.ps1 -Ambiente Teste
```

```powershell
.\rodar_fluxo_funil.ps1 -Ambiente Producao
```

```powershell
.\rodar_fluxo_funil.ps1 -Ambiente Teste -NoSharePoint
```

```powershell
.\rodar_fluxo_funil.ps1 -Ambiente Teste -StartDate 2026-03-01 -EndDate 2026-03-06
```

## 2. Fluxo Comparativo

Objetivo: gerar o JSON comparativo entre 2 periodos e enviar para a pasta final de teste ou producao, de onde o Power Automate publica o card comparativo no Teams.

### O que o fluxo faz

1. reaproveita a mesma base da rodada utilizada no `Insight Funil`
2. calcula o periodo atual e o periodo anterior proporcional
3. gera 2 resumos independentes
4. consolida o JSON comparativo
5. copia o JSON para a pasta sincronizada do SharePoint / Teams
6. o Power Automate publica o card comparativo no Teams

### Arquivos envolvidos

- `rodar_fluxo_comparativo_funil.ps1`
- `gerar_comparativo_resumo_insight_periodos.py`
- `gerar_resumo_insight_json.py`

### Regra operacional

O comparativo deve ser gerado e publicado sempre depois do `Insight Funil`.

### Reaproveitamento real da base

Ate pouco tempo atras, cada fluxo (`rodar_fluxo_funil.ps1` e `rodar_fluxo_comparativo_funil.ps1`) buscava
o "arquivo mais recente" na pasta de extracao de forma **independente**, no momento exato em que cada um
rodava. Isso causava uma divergencia real: se o Power Automate atualizasse a extracao entre uma rodada e
outra (ex.: rodar o Insight as 08:28 e o Comparativo as 08:39), os dois cards do Teams mostravam numeros
diferentes pro "periodo atual", mesmo com o rotulo do periodo identico.

Isso foi corrigido:

1. `rodar_fluxo_funil.ps1`, depois de escolher o DAX1/DAX2 da rodada, grava um arquivo `.ultima_base_funil.json`
   na raiz do projeto com o caminho exato desses dois arquivos e o timestamp da rodada.
2. `rodar_fluxo_comparativo_funil.ps1`, antes de buscar o DAX1/DAX2, tenta ler esse "bilhete":
   - se existir, tiver no maximo 6 horas e os arquivos referenciados ainda existirem, usa exatamente
     os mesmos arquivos que o Insight Funil usou na ultima rodada (aparece no console:
     `Reaproveitando a mesma base do Insight Funil (rodada de ...)`);
   - se nao existir, estiver velho ou os arquivos tiverem sumido, cai no comportamento antigo
     (busca o arquivo mais recente na pasta de extracao).

Na pratica, isso garante que **Insight Funil e Comparativo nascam sempre da mesma base**, desde que o
Comparativo rode logo depois do Insight Funil (dentro da mesma janela de 6h) — o que ja e a ordem de
publicacao recomendada.

### Observacao sobre nao convertidas no comparativo

Quando o comparativo executa o pipeline auxiliar, ele tambem respeita a regra padrao de nao convertidas: `Faturou = 0`, `Aprovado pelo Cliente = 0`, status de recusa/cancelamento e `tipo_perda = todas`. A troca entre `Teste` e `Producao` muda apenas o destino de publicacao; a regra de calculo continua a mesma.

### Regra de mes comercial no comparativo

- o comparativo usa a mesma logica de mes comercial do fluxo principal
- o `Periodo Atual` nasce do mes comercial atual, o `Periodo Anterior` nasce do mes comercial anterior
- a equivalencia entre os periodos e feita pela mesma quantidade de dias uteis observados
- o JSON comparativo carrega `comparativo.periodo_a.label`, `comparativo.periodo_a.dias_uteis`, `comparativo.periodo_b.label`, `comparativo.periodo_b.dias_uteis`, `atualizado_em`
- o campo `mes_comercial` tambem e gerado nos resumos-base e no JSON comparativo, mesmo que nao precise ser exibido no card final

### Regra de dias uteis no comparativo

O script considera automaticamente: sabados, domingos, feriados nacionais do Brasil e Sexta-feira Santa.

### Parametros disponiveis

- `-Ambiente Teste|Producao` — define a pasta sincronizada do SharePoint de destino
- `-StartDate YYYY-MM-DD` — opcional; define o inicio do periodo atual
- `-EndDate YYYY-MM-DD` — opcional; define o fim do periodo atual
- `-NoPublicarComparativo` — gera os arquivos localmente, mas nao copia o comparativo para a pasta sincronizada

### Comportamento padrao

Se `-StartDate` e `-EndDate` nao forem informados:

- inicio = primeiro dia do mes comercial atual, fim = ontem
- `Periodo Atual` = mes comercial atual ate ontem
- `Periodo Anterior` = periodo equivalente dentro do mes comercial anterior, com a mesma quantidade de dias uteis observados no periodo atual

### Comandos

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste
```

```powershell
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Producao
```

```powershell
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste -StartDate 2026-04-01 -EndDate 2026-04-07
```

```powershell
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste -NoPublicarComparativo
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

## 4. Resumo Mensal Historico (auxiliar)

Objetivo: gerar uma planilha mensal historica com os mesmos conceitos de remocao de ruido usados no funil principal, sem depender do Excel consolidado da rodada.

### O que o fluxo faz

1. le os arquivos DAX1 e DAX2 diretamente em `json` ou em base tratada
2. aplica a mesma logica de remocao de ruido do `gerar_oportunidades_reais_codes.py`
3. consolida os indicadores `Sem Ruído` por linha mensal
4. permite gerar a visao por mes comercial ou por mes calendario
5. grava a saida em um Excel historico auxiliar

### Arquivos envolvidos

- `gerar_resumo_mensal_funil.py` (script principal)
- `gerar_resumo_mensal_comercial.py` — nome antigo mantido como atalho para compatibilidade

### Saida principal

- `historico\resumo_mensal_funil_YYYYMMDD_HHMMSS.xlsx`, aba `Resumo_Mensal_Funil`, com as colunas:
  `Mês/Ano Comercial`, `Periodo`, `Enviados_Qtd (Sem Ruído)`, `Enviados_Valor (Sem Ruído)`, `Faturado_Qtd (Sem Ruído)`, `Faturado_Valor (Sem Ruído)`, `Nao_Convertidas_Qtd (Sem Ruído)`, `Nao_Convertidas_Valor (Sem Ruído)`, `Win Rate (Volume) % (Sem Ruído)`, `Win Rate (Valor) % (Sem Ruído)`.

### Modos disponiveis

- `comercial` — padrao do script; cada linha representa um ciclo comercial
- `calendario` — ativado com `--modo calendario`; cada linha representa um mes calendario civil

### Regras importantes

- `--start` e `--end` limitam a janela total da base analisada
- `--modo` define se a agregacao de cada linha sera comercial ou calendario
- quando a base nao cobre o ciclo inteiro, o script mantem a linha e informa no console os meses parciais
- para o primeiro mes comercial sair completo, a extracao deve comecar antes do inicio do mes alvo
  - ex.: para analisar `Jun/25 a Nov/25` em modo comercial, a extracao deve comecar em `01/05/2025`
  - ex.: para analisar `Dez/25 a Abr/26` em modo comercial, a extracao deve comecar em `01/11/2025`

### Exemplos de uso

```powershell
cd C:\analise_funil
py -3.12 .\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --modo calendario --start 2026-04-01 --end 2026-04-30 -o resumo_calendario_abril
```

```powershell
py -3.12 .\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --start 2026-03-31 --end 2026-04-29 -o resumo_comercial_abril
```

Usando os JSONs mais recentes da pasta de extracao:

```powershell
cd C:\analise_funil
$jsonDir = "C:\CAMINHO\PARA\Entrada\FunilExtracao"
$dax1 = Get-ChildItem -LiteralPath $jsonDir -Filter "dax_query_base_orcamentos_funil_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$dax2 = Get-ChildItem -LiteralPath $jsonDir -Filter "dax_query_base_itens_funil_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
py -3.12 .\gerar_resumo_mensal_funil.py -i $dax1.FullName -it $dax2.FullName --modo calendario --start 2025-12-01 --end 2026-04-30 -o resumo_calendario_dez_abr
```

## 5. Estrutura de Pastas

Sugestao de organizacao:

- Entrada bruta: `/Entrada/FunilExtracao`
- Saida insight teste: pasta configurada em `insight_json_dir` do ambiente `Teste`
- Saida insight producao: pasta configurada em `insight_json_dir` do ambiente `Producao`
- Saida comparativo teste: pasta configurada em `comparativo_json_dir` do ambiente `Teste`
- Saida comparativo producao: pasta configurada em `comparativo_json_dir` do ambiente `Producao`

## 6. Regra de Uso

- `rodar_fluxo_funil.ps1` — quando quiser executar a rodada principal do funil e gerar o `Insight Funil`
- `rodar_fluxo_comparativo_funil.ps1` — quando quiser gerar o comparativo depois da rodada principal
- `gerar_resumo_mensal_funil.py` — quando quiser gerar uma visao historica mensal auxiliar da base, em mes comercial ou calendario

Observacao operacional:

- hoje o uso previsto e manual; primeiro rode o `Insight Funil`, depois rode o `Comparativo`
- a extracao base ja deve ter sido gerada pelo Power Automate antes dos comandos locais
- no modo padrao: `Insight Funil` considera mes comercial atual ate ontem; `Comparativo` considera mes comercial atual ate ontem versus periodo equivalente do mes comercial anterior por dias uteis; `Resumo Mensal Funil` usa o modo `comercial`, salvo quando `--modo calendario` for informado

Observacao sobre leitura temporal nos cards:

- no `Insight Funil`, o card mostra: `Mes Comercial`, `Periodo analisado`, `Atualizado em`
- no `Comparativo`, o card mostra: `Periodo Anterior`, `Periodo Atual`, `Atualizado em`

## 7. Resumo Operacional

```powershell
.\rodar_fluxo_funil.ps1 -Ambiente Teste
.\rodar_fluxo_funil.ps1 -Ambiente Producao

.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Producao

.\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --modo calendario --start 2026-04-01 --end 2026-04-30 -o resumo_calendario_abril
.\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --start 2026-03-31 --end 2026-04-29 -o resumo_comercial_abril
```

## 8. Integracao com o Power Automate

O PowerShell do `Insight Funil`:

1. gera `resumo_insight_card_teams_...json`
2. copia esse JSON para a pasta sincronizada correta
3. deixa o Power Automate assumir a postagem no Teams

O PowerShell do `Comparativo`:

1. gera os dois resumos localmente
2. monta `comparativo_resumo_insight_card_teams_...json`
3. copia o comparativo para a pasta sincronizada correta
4. deixa o Power Automate assumir a postagem no Teams

## 9. Publicacao publica (GitHub Pages)

O dashboard interativo (`Dashboard Analise Funil\dashboard_analise_funil.dc.html`, com os cards `Resumo Atual` e `Comparativo`) tem uma versao publicada de verdade, acessivel por link, sem precisar estar na maquina local.

### Como funciona

- Os dois `.ps1` (`rodar_fluxo_funil.ps1` e `rodar_fluxo_comparativo_funil.ps1`) copiam o `.js` gerado a cada rodada (`resumo_insight_dashboard.js` / `comparativo_dashboard.js`) pra pasta configurada em `paths.dashboard_analise_funil_publish_dir` (pasta do SharePoint) — a menos que `-NoPublicarDashboard` seja informado.
- A pasta `docs\` na raiz do repositorio contem a versao publicavel do HTML (copia de `Dashboard Analise Funil\dashboard_analise_funil.dc.html`), hospedada via **GitHub Pages**.
- Em vez de carregar os dados de um arquivo local (`./data/*.js`), o `docs\index.html` aponta os dois `<script src>` pra **URLs de compartilhamento do SharePoint** dos arquivos publicados.
- Isso funciona porque uma tag `<script src>` nao passa pelas travas de CORS que um `fetch()` teria — o navegador carrega o arquivo como um recurso normal, e o cookie de sessao do SharePoint e enviado junto. Resultado: **o link e publico, mas os dados so aparecem pra quem estiver logado na conta Microsoft da empresa** — sem login, o SharePoint bloqueia e nenhum dado carrega.
- Se os dados nao carregarem (falha de login, link expirado, etc.), a pagina mostra um modal explicando que e preciso estar logado, com um botao pra recarregar depois de logar.

### O que isso depende

- Os `.ps1` continuarem publicando os dois `.js` em `dashboard_analise_funil_publish_dir` a cada rodada — sem isso, a pagina publica fica com dado desatualizado.
- Os links de compartilhamento do SharePoint embutidos no `docs\index.html` continuarem validos — se forem revogados/expirarem, e preciso gerar novos links e atualizar os `<script src>` no arquivo.
- O GitHub Pages estar habilitado nas configuracoes do repositorio (Settings → Pages → Branch `main`, pasta `/docs`).

### Limitacao atual

Como `docs\index.html` e uma copia estatica do `.dc.html`, qualquer mudanca visual feita no dashboard local precisa ser replicada manualmente em `docs\index.html` (trocando so os `<script src>` de volta pros links do SharePoint) pra a versao publica acompanhar.

## 10. Publicando uma mudanca de codigo

Depois de editar qualquer arquivo do projeto:

```bash
cd C:\analise_funil
git status
git add nome_do_arquivo
git commit -m "Descricao da mudanca"
git push -u origin main
```

Antes de commitar, vale checar que nao vazou nada sensivel:

```bash
git diff --cached --name-only -z | xargs -0 grep -lI "erico.moraes\|BUNZL"
```

## 11. Observacoes importantes

- hoje a operacao normal nao depende de login manual no Power BI, porque a extracao base vem do Power Automate
- a postagem no Teams hoje acontece via Power Automate
- o fluxo de resumo e o fluxo comparativo podem apontar para canais diferentes no Teams
- o comparativo automatico usa dias uteis automaticamente no modo padrao
- a extracao do Power Automate hoje pode trazer uma janela ampliada para atender resumo e comparativo na mesma rodada, mas o recorte efetivo de cada card e definido pelos scripts locais
- no fim de cada rodada os scripts exibem os principais arquivos gerados
- desde a correcao do `.ultima_base_funil.json`, o Insight Funil e o Comparativo sempre nascem da mesma
  base de extracao (contanto que o Comparativo rode ate 6h depois do Insight Funil)
- a identidade do Git usada nos commits deve ser `Erico Casarano <erico.casarano@exemplo.com>` — se um
  commit sair com outro autor (ex.: e-mail corporativo real), corrija antes de publicar
  (`git commit --amend --author="Erico Casarano <erico.casarano@exemplo.com>"`, so em commits ainda nao
  enviados ao GitHub)
