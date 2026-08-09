# RD Station CRM -> Vendas (SharePoint)

Grava na aba `Vendas` (Tabela4) do `VENDAS_E_CAMPANHA.xlsx` os negocios fechados
a partir de 09/08/2026. Uma linha por produto. Nada anterior ao corte e tocado.

Tudo roda pela interface do GitHub. Nao e preciso Python na sua maquina.

## Arquivos

    sync_rd_vendas.py
    requirements.txt
    .github/workflows/sync.yml

## Modos de execucao

Em Actions -> Run workflow, escolha no menu "O que executar":

| Modo | O que faz |
|---|---|
| listar  | mostra os arquivos da biblioteca com caminho e id |
| dry-run | mostra o que gravaria, sem tocar na planilha |
| gravar  | grava de verdade |

O cron de 15 minutos sempre usa o modo gravar.

## Variaveis (no proprio sync.yml)

| Variavel | Padrao |
|---|---|
| SHAREPOINT_FILE_PATH | /VENDAS_E_CAMPANHA.xlsx |
| CUTOFF_DATE | 2026-08-09 |
| SHEET_NAME | Vendas |
| TABLE_NAME | Tabela4 |
| STAGE_MATCH | FECHAMENTO |
| PIPELINE_IDS | vazio (aceita qualquer funil) |

## Deduplicacao

Chave: Nome | Data de fechamento | Valor Unico | Produtos, aplicada apenas as
linhas a partir do corte. A planilha e a fonte da verdade, nao ha arquivo de
estado. Apagar uma linha faz o script reinseri-la no ciclo seguinte.

Limite: dois negocios identicos do mesmo paciente, mesmo produto, mesmo valor e
mesmo dia sao tratados como um so.

## Colunas L a O

Servico, MM/AAAA ganho, Bonus R$ e Pont. Ranking sao formulas.
O script grava apenas A:K e replica L:O lendo L2:O2 em formato R1C1 a cada
execucao. Alterou a formula na linha 2, as novas linhas seguem a nova versao.

## Nao faca

- Nao renomeie colunas nem a tabela Tabela4
- Nao classifique a aba Vendas durante a execucao
- Nao deixe o arquivo aberto no Excel Desktop com alteracoes nao salvas
