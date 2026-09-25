# Lopex DataCrazy

Painel web para auditoria e rodízio do DataCrazy, preparado para Railway.

## V11 integrado

O botão **EXECUTAR V11** executa a coleta com a mesma lógica do V11: status `unstarted`, `waiting`, `opened`, `automation` e `error`; descarta conversas apagadas, arquivadas, ocultas ou finalizadas; separa sem atendente, FAILED sem atendente e FAILED com atendente; cruza origem Meta/Google e mantém histórico.

O botão **BAIXAR EXCEL V11** gera um arquivo com as abas:
- RESUMO
- TODOS SEM ATENDENTE
- SEM ATENDENTE
- FAILED SEM ATENDENTE
- FAILED COM ATENDENTE
- LEADS UNICOS S ATENDENTE
- FONTES ORIGEM

## Railway

Root Directory: `/lopex-datacrazy-app`
Healthcheck: `/health`
Start: `npm start`

Variáveis opcionais:
- `DATACRAZY_TOKEN` — token privado da API.
- `APP_PASSWORD` — senha do painel.

O token não deve ser salvo no repositório público.
