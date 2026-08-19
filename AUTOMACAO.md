# Automacao do funil

## Visao geral

O projeto hoje possui dois fluxos principais:

1. resumo executivo do funil
2. comparativo de win rate e funil versus o mes anterior

Além deles, o projeto agora possui:

- um script auxiliar para gerar consolidacoes mensais historicas com a mesma logica de remocao de ruido do funil
- um fluxo dedicado ao painel de aprovacao do gestor, que tambem tem uma versao publicada com link
  publico (GitHub Pages) — ver secao "Publicacao publica (GitHub Pages)" mais abaixo

Ambos devem partir da mesma rodada de extracao e seguir a ordem de publicacao abaixo:

1. primeiro publicar o `Insight Funil`
2. depois publicar o `Comparativo`

No desenho atual da automacao local:

- `rodar_fluxo_funil.ps1` e o fluxo principal da rodada
- `rodar_fluxo_comparativo_funil.ps1` gera o comparativo
- `rodar_fluxo_aprovacao_gestor.ps1` atualiza o painel de aprovacao do gestor
- o nome do fluxo principal foi simplificado para refletir melhor esse papel

Modo de operacao atual:

- o Power Automate faz a extracao das consultas DAX
- a execucao local dos scripts acontece manualmente
- o uso normal hoje e por comando em `Teste` ou `Producao`

## Estrategia adotada

A extracao base da rodada hoje vem do Power Automate, com os arquivos JSON gravados na pasta de entrada sincronizada.

Beneficios:

- separa a extracao do processamento local
- permite usar a mesma base para o `Insight Funil` e para o `Comparativo`
- reduz dependencia de login manual no Power BI para a operacao do dia a dia

## Ambientes

Os envios podem ser separados por ambiente:

- `Teste`
- `Producao`

Configuracao atual:

### Producao

- Site SharePoint: `https://bunzlbr.sharepoint.com/sites/InsightsVendas`
- Pasta resumo: `Documentos\Producao\FunilInsights`
- Pasta comparativo: `Documentos\Producao\FunilComparativos`
- Pasta do painel de Aprovacao do Gestor (publicado no GitHub Pages): `Documentos\Producao\StaticWebApp`

### Teste

- Site SharePoint: `https://bunzlbr.sharepoint.com/sites/InsightsVendas`
- Pasta resumo: `Documentos\Teste\FunilInsights_teste`
- Pasta comparativo: `Documentos\Teste\FunilComparativos_teste`

## Fluxo 1. Resumo Executivo do Funil

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
- `tratar_dax1_json_power_automate.py`
- `tratar_dax2_json_power_automate.py`
- `automacao_pipeline.py`
- `gerar_resumo_insight_json.py`

### Saidas principais

- `entrada\dax1_remocao_ruidos_power_automate_YYYYMMDD_HHMMSS.xlsx`
- `entrada\dax2_itens_orcamento_power_automate_YYYYMMDD_HHMMSS.xlsx`
- `historico\oportunidades_reais_auto_YYYYMMDD_HHMMSS.xlsx`
- `historico\itens_perdas_reais_auto_YYYYMMDD_HHMMSS.xlsx`
- `alertas\resumo_insight_power_automate_YYYYMMDD_HHMMSS.json`
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

Opcoes aceitas pelo script de itens nao convertidos:

- `todas`
- `revisadas`
- `sem_revisao`

### Saida detalhada do step 2

O step 2 gera o arquivo `historico\itens_perdas_reais_auto_YYYYMMDD_HHMMSS.xlsx`.

Abas principais:

- `Base_Nao_Convertidas`: orcamentos com `Faturou = 0`, `Aprovado pelo Cliente = 0` e status de recusa/cancelamento, respeitando `tipo_perda`.
- `Itens_Nao_Convertidas`: itens desses orcamentos.
- `Base_Faturados`: orcamentos com `Faturou = 1`, no mesmo recorte de revisao.
- `Itens_Faturados`: itens dos orcamentos faturados.

Rankings mantidos:

- `Ranking_Geral_Final`
- `Ranking_Geral_Revenda`
- `Ranking_Geral_Final_Faturados`
- `Ranking_Geral_Revenda_Faturados`
- `Ranking_Itens_Por_Vendedor`
- `Ranking_Itens_Por_Mes`
- `Ranking_Itens_Vendedor_Mes`
- abas `Top10_<Vendedor>`

Abas removidas para simplificar o arquivo:

- `Ranking_Itens_Geral`
- `Ranking_Itens_Geral_Tipo`
- `Ranking_NC_Recusa_Tipo`
- `Ranking_NC_Recusa_Final`
- `Ranking_NC_Recusa_Revenda`
- `Base_NC_Recusa_Cliente`
- `Itens_NC_Recusa_Cliente`
- `Analise_Preco_Recusa`
- `Ranking_Itens_Tipo_Faturados`

### Precos e classificacao no step 2

Nos rankings, as colunas de preco sao:

- `Preco_Ponderado`: `Valor_Total / Volume_Total`.
- `Preco_Media_Orcamento`: media dos precos calculados por orcamento.
- `Preco_Mediana_Orcamento`: mediana dos precos calculados por orcamento.

Na aba `Analise_Preco_Itens`, as mesmas visoes aparecem separadas para faturados e nao convertidos.

A `Classificacao_Preco` usa a mediana por orcamento como referencia principal:

- `Dif_Preco_Mediana_% >= 3%` e `Win_Rate_Item_% < 50%`: `Possivel sensibilidade a preco`.
- `Dif_Preco_Mediana_% >= 3%`: `Atencao: preco nao convertido maior`.
- `Dif_Preco_Mediana_% <= -3%`: `Nao convertido com preco menor`.
- diferenca entre `-3%` e `+3%`: `Preco similar`.
- menos de 3 orcamentos faturados ou menos de 3 nao convertidos: `Amostra insuficiente`.

### Pipeline tecnico atualmente utilizado

Etapas ativas:

- `step1`: `gerar_oportunidades_reais_codes.py`
- `step2`: `gerar_itens_perdas_reais.py`

Ou seja, hoje a rodada principal esta focada em:

- consolidacao das oportunidades reais
- nao convertidos e rankings de itens
- geracao do `Insight Funil`

### Comando em teste

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_funil.ps1 -Ambiente Teste
```

### Comando em producao

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_funil.ps1 -Ambiente Producao
```

### Comando sem envio ao SharePoint

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_funil.ps1 -Ambiente Teste -NoSharePoint
```

### Comando com periodo especifico

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_funil.ps1 -Ambiente Teste -StartDate 2026-03-01 -EndDate 2026-03-06
```

### Parametros disponiveis no fluxo principal atual

- `-Ambiente Teste|Producao`
  - define a pasta sincronizada do SharePoint de destino
- `-StartDate YYYY-MM-DD`
  - define a data inicial da extracao
- `-EndDate YYYY-MM-DD`
  - define a data final da extracao
- `-NoSharePoint`
  - gera o JSON final apenas localmente, sem copiar para a pasta sincronizada

### Comportamento padrao do fluxo principal atual

Se `-StartDate` e `-EndDate` nao forem informados:

- o script usa os arquivos JSON mais recentes disponiveis na pasta de extracao
- inicio = primeiro dia do mes comercial atual
- fim = ontem
- o `Insight Funil` e gerado somente com esse recorte
- mesmo que a extracao contenha uma janela mais ampla para atender o comparativo, o card principal nao deve considerar esse intervalo adicional

### Regra de mes comercial no fluxo principal

- o mes comercial fecha no penultimo dia util do mes
- o dia seguinte ao fechamento passa a pertencer ao mes comercial seguinte
- o JSON do `Insight Funil` hoje carrega:
  - `titulo`
  - `mes_comercial`
  - `periodo`
  - `atualizado_em`
- isso permite ao Adaptive Card diferenciar claramente:
  - o ciclo comercial da rodada
  - a parcela efetivamente analisada ate ontem

## Fluxo 2. Comparativo

### O que o fluxo faz

1. reaproveita a mesma base da rodada utilizada no `Insight Funil`
2. calcula o periodo atual e o periodo anterior proporcional
3. gera 2 resumos independentes
4. consolida o JSON comparativo
5. copia o JSON para a pasta sincronizada do SharePoint / Teams
6. o Power Automate publica o card comparativo no Teams

### Arquivos envolvidos

- `rodar_fluxo_comparativo_funil.ps1`
- `gerar_comparativo_win_rate_periodos.py`
- `gerar_resumo_insight_json.py`

### Regra operacional

- o comparativo deve ser gerado depois do `Insight Funil`
- a publicacao do comparativo deve ocorrer sempre depois da publicacao do `Insight Funil`

### Reaproveitamento real da base (corrigido)

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

### Comando padrao em teste

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste
```

### Comando padrao em producao

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Producao
```

### Comando com periodo explicito

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste -StartDate 2026-04-01 -EndDate 2026-04-07
```

### Comando sem publicar no SharePoint

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_comparativo_funil.ps1 -Ambiente Teste -NoPublicarComparativo
```

### Variaveis disponiveis no comparativo automatico

- `-Ambiente Teste|Producao`
  - define a pasta sincronizada do SharePoint de destino
- `-StartDate YYYY-MM-DD`
  - opcional; define o inicio do periodo atual
- `-EndDate YYYY-MM-DD`
  - opcional; define o fim do periodo atual
- `-NoPublicarComparativo`
  - gera os arquivos localmente, mas nao copia o comparativo para a pasta sincronizada

### Comportamento padrao do comparativo automatico

Se `-StartDate` e `-EndDate` nao forem informados:

- inicio = primeiro dia do mes comercial atual
- fim = ontem
- `Periodo Atual` = mes comercial atual ate ontem
- `Periodo Anterior` = periodo equivalente dentro do mes comercial anterior
- ajuste = mesma quantidade de dias uteis observados no periodo atual

## Fluxo 3. Painel de Aprovacao do Gestor

### O que o fluxo faz

1. le os JSONs mais recentes da pasta sincronizada `/Entrada/FunilExtracao` (Bloco 1/DAX2) e, se `-UsarBloco3`
   for usado, da pasta `/Entrada/OrcamentosRevisadosExtracao` (Bloco 2/Bloco 3)
2. converte os JSONs brutos para arquivos tratados em `entrada/`
3. gera ou reaproveita a base de oportunidades
4. monta a analise de aprovacao do gestor em Excel e JSON
5. atualiza o arquivo do slide HTML interativo local
6. publica a copia do slide na pasta do SharePoint (`Producao\StaticWebApp`), que alimenta a versao
   publica no GitHub Pages (ver secao "Publicacao publica" mais abaixo)

### Arquivos envolvidos

- `rodar_fluxo_aprovacao_gestor.ps1`
- `tratar_dax1_json_power_automate.py`
- `tratar_dax2_json_power_automate.py`
- `automacao_pipeline.py`
- `gerar_analise_aprovacao_gestor.py`
- `atualizar_slide_aprovacao_gestor_json.py`

### Blocos de extracao (`-UsarBloco3`)

O painel pode usar duas fontes diferentes pra base "com ruido" (DAX1):

- **sem `-UsarBloco3`** (padrao): usa `Entrada\FunilExtracao\dax1_funil_powerbi_*.json` — a mesma extracao
  do fluxo principal do funil.
- **com `-UsarBloco3`**: usa `Entrada\OrcamentosRevisadosExtracao\bloco3_aprovacao_cliente_*.json` no lugar
  do DAX1, e adicionalmente le `Entrada\OrcamentosRevisadosExtracao\bloco2_revisao_gestor_*.json` (usado
  so pra calcular os KPIs de revisao: `qtd_orcamentos_com_revisao`, `qtd_revisoes_total`,
  `media_revisoes_por_orcamento`, e o total de revisoes por vendedor na tabela executiva).

> **Nao confirmado oficialmente**: pelo nome dos arquivos e como o codigo usa cada um, "Bloco 2" parece ser
> os dados de revisao do gestor e "Bloco 3" os dados pos-aprovacao do cliente — mas isso foi inferido lendo
> o codigo e os arquivos, nao documentado por quem desenhou esse fluxo no Power Automate. Nao existe
> "Bloco 1" no codigo (o nome nao aparece em nenhum parametro). Vale confirmar com quem montou a extracao
> `OrcamentosRevisadosExtracao` o que cada bloco representa de fato. O DAX2 (itens) sempre vem de
> `FunilExtracao`, independente de `-UsarBloco3`.

O comando usado hoje no dia a dia e:

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_aprovacao_gestor.ps1 -UsarBloco3 -Visao com_ruido
```

### Regra de visao

- `com_ruido`
  - agora e a visao padrao do painel
  - indicada para leitura de esforco da gestao
  - nao gera `oportunidades_reais_auto_*.xlsx` automaticamente quando o comando roda no modo padrao
- `sem_ruido`
  - usa a base de oportunidades reais sem ruido
  - se `-Oportunidades` nao for informado, o proprio script gera o `oportunidades_reais_auto_*.xlsx` automaticamente antes da analise

### Saidas principais

- `entrada\dax1_remocao_ruidos_aprovacao_gestor_YYYYMMDD_HHMMSS.xlsx`
- `entrada\dax2_itens_orcamento_aprovacao_gestor_YYYYMMDD_HHMMSS.xlsx`
- `entrada\bloco2_revisao_gestor_aprovacao_gestor_YYYYMMDD_HHMMSS.xlsx` (so com `-UsarBloco3`)
- `historico\analise_aprovacao_gestor_YYYYMMDD_HHMMSS.xlsx`
- `alertas\analise_aprovacao_gestor_YYYYMMDD_HHMMSS.json`
- `Slide de Win Rate Interativo\data\aprovacao_gestor_latest.js` (versao local)
- copia publicada em `paths.aprovacao_gestor_publish_dir` (pasta do SharePoint), a menos que `-NoPublicar`
  seja informado

### Parametros do painel

- `-UsarBloco3`
  - troca a fonte do DAX1 pelo Bloco 3 e ativa a leitura do Bloco 2 (ver secao acima)
- `-Visao com_ruido|sem_ruido`
  - `com_ruido` (padrao): nao gera oportunidades reais automaticamente
  - `sem_ruido`: gera (ou reaproveita, via `-Oportunidades`) a base sem ruido
- `-Oportunidades <caminho>`
  - usa um arquivo de oportunidades reais ja pronto, em vez de gerar um novo
- `-NoPublicar`
  - atualiza o slide so localmente, sem copiar pra pasta do SharePoint
- `-JsonDir` / `-Bloco3JsonDir`
  - opcionais; se omitidos, usam `paths.json_dir` / `paths.bloco3_json_dir` do `automacao_config.json`
- `-ForceRun`
  - ignora a trava de duplicidade (usa quando quiser forcar uma rodada nova mesmo tendo rodado a
    mesma combinacao de visao/oportunidades/bloco3 ha poucos minutos)

### Configuracao necessaria (`automacao_config.json`)

Esses caminhos nao ficam mais fixos no `.ps1` (eram valores com usuario/pasta local hardcoded antes) —
agora vem do `automacao_config.json`, em `paths`:

```json
"paths": {
  "history_dir": "./historico",
  "json_dir": "C:\\CAMINHO\\PARA\\Entrada\\FunilExtracao",
  "bloco3_json_dir": "C:\\CAMINHO\\PARA\\Entrada\\OrcamentosRevisadosExtracao",
  "aprovacao_gestor_publish_dir": "C:\\CAMINHO\\PARA\\Producao\\StaticWebApp"
}
```

O `automacao_config.example.json` (versionado no Git) tem essa mesma estrutura com placeholders, pra
servir de modelo caso o `automacao_config.json` real precise ser recriado.

### Comando padrao com ruido

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_aprovacao_gestor.ps1
```

Observacao:

- esse comando agora e o caminho mais rapido para atualizar o painel em `com_ruido`
- ele nao dispara a geracao automatica de `oportunidades_reais_auto`

### Comando explicito com ruido

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_aprovacao_gestor.ps1 -Visao com_ruido
```

### Comando sem ruido

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_aprovacao_gestor.ps1 -Visao sem_ruido
```

### Comando sem ruido com arquivo explicito de oportunidades

```powershell
cd C:\analise_funil
Set-ExecutionPolicy -Scope Process Bypass
.\rodar_fluxo_aprovacao_gestor.ps1 -Visao sem_ruido -Oportunidades .\historico\oportunidades_reais_auto_20260816_101500.xlsx
```

### Regra de mes comercial no comparativo

- o comparativo usa a mesma logica de mes comercial do fluxo principal
- o `Periodo Atual` nasce do mes comercial atual
- o `Periodo Anterior` nasce do mes comercial anterior
- a equivalencia entre os periodos e feita pela mesma quantidade de dias uteis observados
- o JSON comparativo hoje carrega:
  - `comparativo.periodo_a.label`
  - `comparativo.periodo_a.dias_uteis`
  - `comparativo.periodo_b.label`
  - `comparativo.periodo_b.dias_uteis`
  - `atualizado_em`
- o campo `mes_comercial` tambem e gerado nos resumos-base e no JSON comparativo, mesmo que nao precise ser exibido no card final

### Regra de dias uteis no comparativo

O script considera automaticamente:

- sabados
- domingos
- feriados nacionais do Brasil
- Sexta-feira Santa

## Publicacao publica (GitHub Pages)

Alem da copia local (`Slide de Win Rate Interativo\`), o painel de **Aprovacao do Gestor** tem uma versao
publicada de verdade, acessivel por link, sem precisar estar na maquina local.

### Como funciona

- O repositorio (`https://github.com/ericocasarano/trilha-aprovacao-orcamentos-labor`) e **publico** no
  GitHub, com o codigo-fonte do projeto (scripts, HTML, etc.) — mas **sem** os dados gerados a cada
  rodada nem segredos (isso fica de fora via `.gitignore`).
- A pasta `docs/` na raiz do repositorio contem a versao publicavel do HTML (copia de
  `Slide de Win Rate Interativo\Aprovacao Gestor Labor.dc.html`), hospedada via **GitHub Pages** em:

  `https://ericocasarano.github.io/trilha-aprovacao-orcamentos-labor/`

- Em vez de carregar os dados de um arquivo local (`./data/aprovacao_gestor_latest.js`), o `docs/index.html`
  aponta o `<script src>` pra uma **URL de compartilhamento do SharePoint** do arquivo
  `aprovacao_gestor_latest.js` publicado em `Producao\StaticWebApp`.
- Isso funciona porque uma tag `<script src>` nao passa pelas travas de CORS que um `fetch()` teria — o
  navegador carrega o arquivo como um recurso normal (igual um `<img>` ou CDN), e o cookie de sessao do
  SharePoint e enviado junto. Resultado: **o link e publico, mas os dados so aparecem pra quem estiver
  logado na conta Microsoft da empresa** — sem login, o SharePoint bloqueia e nenhum dado carrega.
- Se os dados nao carregarem (falha de login, link expirado, etc.), a pagina mostra um modal explicando
  que e preciso estar logado, com um botao pra recarregar depois de logar.

### O que isso depende

- O `.ps1` continuar publicando `aprovacao_gestor_latest.js` em `Producao\StaticWebApp` a cada rodada
  (ver secao do Fluxo 3 acima) — sem isso, a pagina publica fica com dado desatualizado.
- O link de compartilhamento do SharePoint embutido no `docs/index.html` continuar valido — se ele for
  revogado/expirar, e preciso gerar um novo link e atualizar o `<script src>` no arquivo.
- O GitHub Pages estar habilitado nas configuracoes do repositorio (Settings → Pages → Branch `main`,
  pasta `/docs`).

### Limitacao atual

Hoje **so o painel de Aprovacao do Gestor** tem essa publicacao. Os outros dois HTMLs
(`Win Rate Labor.dc.html`, usado pelo Resumo Atual e Comparativo) continuam carregando os dados de forma
local (`./data/comparativo_latest.js` e `./data/resumo_atual_latest.js`) — nenhum dos dois e copiado pro
SharePoint nem publicado no GitHub Pages ainda. Pra ter o mesmo link publico pra esses dois, seria preciso
replicar o mesmo padrao (passo de copia no `.ps1` + copia do HTML em `docs/` + link de compartilhamento
do SharePoint).

### Publicando uma mudanca de codigo

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

## Fluxo 4. Resumo Mensal Historico

### O que o fluxo faz

1. le os arquivos DAX1 e DAX2 diretamente em `json` ou em base tratada
2. aplica a mesma logica de remocao de ruido do `gerar_oportunidades_reais_codes.py`
3. consolida os indicadores `Sem Ruído` por linha mensal
4. permite gerar a visao por mes comercial ou por mes calendario
5. grava a saida em um Excel historico auxiliar

### Arquivos envolvidos

- `gerar_resumo_mensal_funil.py`
- `gerar_resumo_mensal_comercial.py`
  - nome antigo mantido como atalho para compatibilidade

### Saida principal

- `historico\resumo_mensal_funil_YYYYMMDD_HHMMSS.xlsx`

### Estrutura da planilha

A planilha gerada possui a aba `Resumo_Mensal_Funil` com as colunas:

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

### Modos disponiveis

- modo `comercial`
  - padrao do script
  - cada linha representa um ciclo comercial
- modo `calendario`
  - ativado com `--modo calendario`
  - cada linha representa um mes calendario civil

### Regras importantes

- `--start` e `--end` limitam a janela total da base analisada
- `--modo` define se a agregacao de cada linha sera comercial ou calendario
- quando a base nao cobre o ciclo inteiro, o script mantem a linha e informa no console os meses parciais
- para o primeiro mes comercial sair completo, a extracao deve comecar antes do inicio do mes alvo
  - exemplo:
    - para analisar `Jun/25 a Nov/25` em modo comercial, a extracao deve comecar em `01/05/2025`
    - para analisar `Dez/25 a Abr/26` em modo comercial, a extracao deve comecar em `01/11/2025`

### Exemplos de uso

Mes calendario:

```powershell
cd C:\analise_funil
py -3.12 .\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --modo calendario --start 2026-04-01 --end 2026-04-30 -o resumo_calendario_abril
```

Mes comercial:

```powershell
cd C:\analise_funil
py -3.12 .\gerar_resumo_mensal_funil.py -i ".\entrada\dax1_jan_maio.xlsx" -it ".\entrada\dax2_jan_maio.xlsx" --start 2026-03-31 --end 2026-04-29 -o resumo_comercial_abril
```

Usando os JSONs mais recentes da pasta de extracao:

```powershell
cd C:\analise_funil
$jsonDir = "C:\CAMINHO\PARA\Entrada\FunilExtracao"
$dax1 = Get-ChildItem -LiteralPath $jsonDir -Filter "dax1_funil_powerbi_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$dax2 = Get-ChildItem -LiteralPath $jsonDir -Filter "dax2_itens_powerbi_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
py -3.12 .\gerar_resumo_mensal_funil.py -i $dax1.FullName -it $dax2.FullName --modo calendario --start 2025-12-01 --end 2026-04-30 -o resumo_calendario_dez_abr
```

## Integracao com o Power Automate

### Resumo

O PowerShell:

1. gera `resumo_insight_...json`
2. copia esse JSON para a pasta sincronizada correta
3. deixa o Power Automate assumir a postagem no Teams

### Comparativo

O comparativo:

1. gera os dois resumos localmente
2. monta `comparativo_win_rate_periodos_...json`
3. copia o comparativo para a pasta sincronizada correta
4. deixa o Power Automate assumir a postagem no Teams

## Observacoes importantes

- hoje a operacao normal nao depende de login manual no Power BI, porque a extracao base vem do Power Automate
- a postagem no Teams hoje acontece via Power Automate
- o fluxo de resumo e o fluxo comparativo podem apontar para canais diferentes no Teams
- o comparativo automatico usa dias uteis automaticamente no modo padrao
- a extracao do Power Automate hoje pode trazer uma janela ampliada para atender resumo e comparativo na mesma rodada, mas o recorte efetivo de cada card e definido pelos scripts locais
- o win rate principal continua usando a data de criacao do orcamento
- o pipeline agora tambem calcula um win rate adicional por data efetiva de faturamento: usa `Data de Faturamento` quando existir e, quando estiver vazia, usa `Data de Criacao`
- o arquivo de saida agora inclui a aba `Lista_WR_Data_Fat`, com os orcamentos considerados nesse segundo recorte
- a aba `Comparativo_Geral_Total` continua exibindo apenas a leitura original por data de criacao
- a nova leitura consolidada fica separada na aba `Comp_Geral_Total_Data_Fat`
- no `Insight Funil`, a exibicao recomendada no card e:
  - `Mes Comercial`
  - `Periodo analisado`
  - `Atualizado em`
- no `Comparativo`, a exibicao recomendada no card e:
  - `Periodo Anterior`
  - `Periodo Atual`
  - `Atualizado em`
- no fim de cada rodada os scripts exibem os principais arquivos gerados
- desde a correcao do `.ultima_base_funil.json`, o Insight Funil e o Comparativo sempre nascem da mesma
  base de extracao (contanto que o Comparativo rode ate 6h depois do Insight Funil)
- o painel de Aprovacao do Gestor tem link publico (GitHub Pages), gated por login do SharePoint; os
  outros dois paineis (Resumo Atual e Comparativo) ainda so funcionam localmente
- a identidade do Git usada nos commits deve ser `Erico Casarano <erico.casarano@exemplo.com>` — se um
  commit sair com outro autor (ex.: e-mail corporativo real), corrija antes de publicar
  (`git commit --amend --author="Erico Casarano <erico.casarano@exemplo.com>"`, so em commits ainda nao
  enviados ao GitHub)
