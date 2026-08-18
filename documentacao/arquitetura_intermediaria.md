# Arquitetura intermediaria para automacao dos alertas

## Objetivo

Registrar uma alternativa intermediaria para evoluir a automacao atual sem perder o que ja foi construido.

A ideia e reduzir a dependencia de login no Power BI pelo PowerShell, mantendo temporariamente a logica analitica em Python ate definirmos o executor definitivo com TI.

## Cenario atual

Hoje o fluxo roda localmente em `C:\analise_funil` e faz:

1. PowerShell autentica no Power BI.
2. PowerShell executa consultas DAX via API do Power BI.
3. Arquivos extraidos sao tratados em Python.
4. Pipeline Python gera as analises.
5. Python gera o JSON do alerta.
6. PowerShell copia o JSON para uma pasta sincronizada do SharePoint.
7. Power Automate detecta o JSON.
8. AI Builder gera o insight.
9. Teams recebe o Adaptive Card.

## Problema do modelo atual

O fluxo depende de:

- login manual no Power BI pelo PowerShell
- sessao autenticada do usuario
- maquina local ligada
- OneDrive/SharePoint sincronizado

Isso dificulta agendamento confiavel e uso em producao.

## Proposta intermediaria

Mover a extracao do Power BI para o Power Automate, mantendo o Python como motor analitico.

Fluxo proposto:

1. Power Automate roda em agenda.
2. Power Automate usa conector do Power BI para executar as consultas DAX.
3. Power Automate salva os resultados da DAX 1 e DAX 2 no SharePoint.
4. Python/PowerShell deixa de acessar a API do Power BI.
5. Python passa a ler os arquivos extraidos do SharePoint/pasta sincronizada.
6. Python executa as analises e gera o JSON final.
7. O JSON final e salvo no SharePoint.
8. Power Automate publica no Teams via AI Builder + Adaptive Card.

## O que essa arquitetura resolve

- Remove a necessidade de login no Power BI dentro do PowerShell.
- Usa a conexao autenticada do Power Automate com o Power BI.
- Reduz o risco de falhas por `Login first with Login-PowerBIServiceAccount`.
- Mantem a logica Python atual com menor retrabalho.

## O que essa arquitetura nao resolve sozinha

- Ainda precisa de um executor para rodar Python.
- Se o Python continuar na maquina local, a maquina ainda precisa estar ligada.
- Ainda depende de OneDrive/SharePoint sincronizado se o processamento continuar local.

## Executor necessario para o Python

Opcoes possiveis:

- maquina local, apenas como etapa temporaria
- servidor/VM
- n8n no servidor existente
- Azure Automation
- Azure Function
- Fabric Notebook/Pipeline, se houver capacidade e permissao

## Desenho simplificado

```text
Power Automate agendado
> Executa DAX no Power BI
> Salva DAX1 e DAX2 no SharePoint
> Executor roda Python lendo arquivos do SharePoint
> Python gera JSON de alerta
> JSON vai para SharePoint
> Power Automate + AI Builder
> Teams Adaptive Card
```

## Comparacao com o modelo atual

### Modelo atual

```text
PowerShell local
> Login Power BI
> API Power BI
> Python
> JSON
> SharePoint
> Power Automate
> Teams
```

### Modelo intermediario

```text
Power Automate
> Power BI
> SharePoint
> Python
> JSON
> SharePoint
> Power Automate
> Teams
```

## Pontos para validar com TI

- Se o Power Automate pode executar as consultas DAX necessarias no dataset.
- Se ha limite de volume/tempo para a extracao via conector Power BI.
- Onde o Python revisado deve rodar.
- Se n8n pode atuar como executor/orquestrador.
- Se existe alternativa Microsoft aprovada, como Azure Automation ou Fabric.
- Se a arquitetura final deve eliminar dependencia de OneDrive sincronizado.

## Observacao

Essa arquitetura e uma etapa intermediaria. Ela nao substitui a necessidade de um ambiente estavel para execucao do Python, mas pode reduzir a dependencia de login manual no Power BI enquanto a solucao definitiva e definida.
