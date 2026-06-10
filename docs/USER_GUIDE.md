# Guia de utilizacao

## Comecar no OrderFlow

1. Abrir `http://127.0.0.1:8030`.
2. Fazer login.
3. Criar ou abrir uma work session.
4. Escolher o cliente.
5. Carregar email `.eml`, rever encomendas e usar os exports necessarios.

## Abrir Shortage Control

O botao `Shortages` fica no topo do OrderFlow. A ferramenta abre em `/shortages/`.

## Criar uma sessao de shortages

1. Criar ou abrir uma work session na pagina inicial do OrderFlow.
2. Abrir `/shortages/`.
3. Abrir a mesma sessao na lista `OrderFlow sessions`.
4. Carregar uma encomenda `.eml` ou um ficheiro Excel/CSV.

Nao e preciso criar uma segunda sessao em Shortages. Os shortages ficam ligados ao mesmo `sessionId` da work session do OrderFlow.

## Controlar shortages

1. Abrir `History`.
2. Escolher a data.
3. Abrir a lista do cliente/ponto de entrega.
4. Editar a coluna `Shortage`.
5. Clicar em `Update delivery point`.

## Exportar Manco's

1. Abrir uma lista guardada no historico.
2. Preencher pelo menos uma linha com shortage.
3. Clicar em `Export Manco's`.
