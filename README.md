# OrderFlow

OrderFlow combina o fluxo completo de encomendas do antigo GreenOps com uma ferramenta de Shortage Control integrada, usando a mesma UI verde e limpa do projeto Shortages.

## Como correr localmente

```powershell
cd C:\Users\Beatriz\Desktop\Project\OrderFlow
py -m pip install -r requirements.txt
py .\app.py
```

Abrir:

```text
http://127.0.0.1:8030
```

## Paginas principais

- `/` - OrderFlow principal: sessoes, clientes, orders, exports, CMR, Leverschema e Laadschema.
- `/shortages/` - ferramenta Shortage Control dentro do OrderFlow.

## Fluxo esperado

1. Entrar no OrderFlow.
2. Criar ou abrir uma work session.
3. Usar os clientes e ferramentas principais para processar encomendas.
4. Abrir `Shortages` quando for preciso controlar mancos/shortages.
5. Dentro de Shortages, abrir a mesma work session e carregar uma encomenda `.eml` ou ficheiro Excel/CSV.

Quando um `.eml` e carregado em Shortages, o parser de encomendas do OrderFlow transforma automaticamente a encomenda em linhas de shortage guardadas na sessao ativa.
Nao e criada uma sessao separada em Shortages: os registos ficam associados ao mesmo `workSessionId`.

## Estrutura

- `app.py` - servidor local combinado, porta `8030`.
- `public/index.html`, `public/app.js`, `public/styles.css` - app principal OrderFlow.
- `public/shortages/` - ferramenta Shortage Control.
- `greenops_shortage_bridge.py` - ponte entre o preview de encomendas e o modelo Shortages.
- `api/` - endpoints serverless para Vercel.
- `email_order_app/` - parsers e writers de encomendas.
- `CMR/` e templates Excel - ficheiros usados pelos exports.
- `data/` - persistencia local quando nao ha Upstash/KV.

## Persistencia

Localmente, os dados ficam em `data/`.

Em deploy, pode usar Upstash/Vercel KV com:

- `UPSTASH_REDIS_REST_URL`
- `UPSTASH_REDIS_REST_TOKEN`

ou:

- `KV_REST_API_URL`
- `KV_REST_API_TOKEN`
