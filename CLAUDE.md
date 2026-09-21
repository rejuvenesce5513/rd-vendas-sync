# rd-vendas-sync — Clínica Rejuvenesce

Script Python que alimenta as planilhas do SharePoint a partir do RD Station CRM e do
Feegow. Roda por GitHub Actions, disparado a cada 15 minutos pelo cron-job.org.

Arquivos: `sync_rd_vendas.py` e `.github/workflows/sync.yml`.

## Antes de mexer

1. O `dry-run` prova que o **cálculo** está certo, nunca que a **gravação** chega ao
   destino. Já aconteceu duas vezes: coluna nova aparecendo no log e gravando nada.
   Depois do primeiro `gravar`, abra a planilha e confira.
2. Toda escrita passa por `mesclar()`. Gravar um intervalo manda o intervalo inteiro —
   sem mesclar, as colunas que este script não preenche são apagadas.
3. Célula preenchida nunca é sobrescrita com vazio.

## O que sincroniza

| Destino | Origem | Casamento |
|---|---|---|
| `Vendas` (Tabela4, A:O) | RD Station, funil comercial | ID RD, depois nome+data+produto |
| `Prevendas` (A:V) | RD Station, funil pré-vendas | ID RD |
| `Marcações` (A:O) | Feegow, agenda | ID Agendamento |
| `Cirurgias.xlsx` (A:F) | Feegow, procedimento 8 | ID Agendamento |

## Modos

`workflow_dispatch` com o parâmetro `modo`:

```
gravar            ciclo completo, o que o cron dispara
dry-run           mostra o que faria, sem escrever
fg-procedimentos  lista os procedimentos do Feegow com id
fg-agenda-dry     diagnóstico da aba Marcações
fg-agenda         grava só a aba Marcações
cirurgias-diag    distribuição por procedimento e campos disponíveis
cirurgias-dry     diagnóstico da planilha de cirurgias
cirurgias         grava só as cirurgias
feegow            cobertura do ID Feegow, no RD e na planilha
feegow-bf-dry     retroativo do ID Feegow, sem escrever
feegow-bf         retroativo do ID Feegow
pipelines         funis do RD
campos-pv         campos personalizados do funil de pré-vendas
bruto-pv          exemplo cru de negociação
pv-dry            diagnóstico da aba Prevendas
prevendas         grava só a aba Prevendas
```

## Segredos

`RD_TOKEN`, `FEEGOW_TOKEN`, `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`,
`AZURE_CLIENT_SECRET`, `SHAREPOINT_DRIVE_ID`. Nunca no código.

## Particularidades das APIs

### RD Station

O filtro por funil não funciona na listagem — a API não devolve `deal_pipeline_id`.
Identificamos o funil de pré-vendas pelo **nome da etapa**.

O primeiro contato só existe em `deal_stage_histories`, que exige buscar cada negócio
individualmente. Teto de 80 por ciclo.

O ID do contato exige `GET /deals/{id}/contacts`. Alguns negócios não têm contato
vinculado e retornam lista vazia — não é erro.

Nome da negociação vem com o prefixo `Pré-vendas - Avaliação gratuita -`. Quando o
resto está vazio ou é apelido curto, usamos o nome do contato — **mas só se a inicial
bater**, senão trocamos pessoas. Já quase aconteceu.

### Feegow

`appoints/search` recusa intervalos de 6 meses ou mais. `fg_janelas()` divide em blocos
de 150 dias.

A paginação às vezes repete um agendamento. Deduplicamos por `agendamento_id`.

O endpoint de paciente não está na documentação pública. `fg_nome_paciente()` tenta
`/patient/search`, `/patient/informations`, `/patient/information` e `/patient/list`,
e memoriza o que responder. Hoje funciona o primeiro.

Valores vêm como texto no formato `R$ 1.500,50`. `toNumF()` converte.

**Não existe campo de "realizado em".** Gravamos o momento em que o sync viu o status
virar `Atendido`, e nunca reescrevemos depois. Precisão de 15 minutos.

Procedimento de cirurgia: **id 8**. Avaliações: 1, 3 e 16.

## Janelas

| Sincronização | Janela |
|---|---|
| Vendas | fechamento a partir de `CUTOFF_DATE` (2026-08-01) |
| Marcações | hoje − 30 até hoje + 90 |
| Cirurgias | hoje − 30 até hoje + 180, móvel |

A janela de cirurgias é móvel de propósito: data fixa expira na virada do ano e para de
trazer registros em silêncio.

## Teto por ciclo

`MAX_UPD` limita as atualizações por execução. Padrão 60. Subir para 200 acelera uma
carga inicial, mas um ciclo chegou a 35 minutos — e o cron dispara a cada 15. Volte para
60 depois de drenar.

## Calendário de manutenção

- **10/08/2027** — o token do GitHub usado pelo cron-job.org expira
- **a cada 60 dias** — um commit no repositório, senão o GitHub desativa o agendamento
- **segredo do Azure** — conferir a validade em Certificados e segredos

## Problemas de dados conhecidos

- Data de nascimento gravada na coluna do ID do paciente, em algumas linhas
- Nome de paciente com texto de campanha: `⭐ Envie esta mensagem e ...`
- Fonte do lead vazia em cerca de 37% dos registros
- Quatro pacientes com duas negociações ganhas e o mesmo ID Feegow — pode ser venda em
  duas partes ou duplicidade de CRM; não foi apurado
- Aba `Campanha` com `#REF!`, não tratada
