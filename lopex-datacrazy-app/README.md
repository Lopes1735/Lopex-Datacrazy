# Lopex DataCrazy

Painel web para o fluxo que antes era feito pelos BATs:

- coleta de conversas atuais sem atendente usando a mesma lógica da V11;
- FAILED separado e com precedência sobre status normais;
- cruzamento de origem por telefone nas planilhas de captação Google Ads / Meta Ads;
- histórico de mensagens por Conversation ID;
- validação manual por checkbox;
- rodízio entre 9 consultores, excluindo Doroty e Patricia;
- exportação XLSX com origem e plano do rodízio.

## Variáveis Railway

- `APP_PASSWORD` (recomendado): protege o painel.
- `DATACRAZY_TOKEN` (opcional): se não definir, o painel pede a chave por sessão.

Nenhum token está salvo no repositório.

## Rodízio

O painel gera o plano e valida os IDs dos consultores na API. A transferência automática dentro do DataCrazy permanece desativada até a rota interna de **Transferir Atendente** ser confirmada.

## Execução

```bash
npm install
npm start
```

Healthcheck: `/health`
