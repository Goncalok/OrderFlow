# Arquitetura

## Visao geral

OrderFlow tem uma pagina principal de operacao de encomendas e uma ferramenta de Shortage Control integrada.

```mermaid
flowchart LR
  A["OrderFlow /"] --> B["Work sessions"]
  B --> C["Clients"]
  C --> D["Orders, exports, CMR"]
  C --> E["Leverschema"]
  C --> F["Laadschema"]
  A --> G["Shortages tool /shortages/"]
  G --> H["Daily shortage sessions"]
```

## Ingest de encomenda para Shortages

```mermaid
flowchart LR
  A["Work session criada em /"] --> B["/shortages/ abre a mesma sessao"]
  B --> C["Upload .eml em /shortages/"]
  C --> D["/api/orders/ingest"]
  D --> E["parse_uploaded_email"]
  E --> F["greenops_shortage_bridge.py"]
  F --> G["Preview Shortages com workSessionId"]
  G --> H["Sessao guardada"]
```

Cada encomenda encontrada no `.eml` vira uma entrada de shortage separada:

- `orderedQuantity` vem da quantidade encomendada.
- `deliveredQuantity` comeca igual ao ordered.
- `shortageQuantity` comeca em `0`.
- O utilizador atualiza os shortages reais mais tarde no historico.
- `workSessionId` liga os shortages a mesma sessao criada na pagina inicial.

## UI

O CSS principal em `public/styles.css` aplica a linguagem visual do Shortages ao OrderFlow:

- fundo verde claro
- topo verde escuro
- paineis brancos
- bordas de 8px
- botoes verdes
- tabelas e formularios consistentes

`public/shortages/styles.css` mantem a mesma base visual para a ferramenta de shortages.
