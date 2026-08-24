# Precificador — Luz Sollar

Painel web para ajustar a precificação das lojas de marketplace a partir da planilha
**Cópia de Luz sollar** (Google Sheets). Site estático, um único arquivo, sem build e sem servidor.

👉 **Abrir:** basta abrir o `index.html` no navegador, ou publicar no GitHub Pages.

## O que dá para fazer

- **8 lojas** em um seletor: Luz Sollar, E-Steck, E-Brilia, E-Lumin, Apollo, Central Lux, Shopee e TikTok.
- **Variáveis da loja** editáveis: NF/Imposto, ADS, Lucro alvo, Devolução e comissões (Clássico / Premium).
- **Ajuste por produto**: custo, desconto, frete/taxa fixa, promoção ativa e rebate — ou travar
  manualmente o preço mínimo, o preço de venda e o preço original.
- **Catálogo de custos** (aba *Custo de produtos*): alterar o custo de um SKU propaga para os kits
  (KIT 5, KIT 10…) e para todas as lojas.
- **Pesquisa de mercado**: média e os 5 preços concorrentes, com comparação contra o preço de venda.
- **Filtros**: com promoção ativa, margem negativa, preço acima do mercado, linhas editadas, sem custo.
- **Exportar CSV** separado por tabulação, com a coluna `Linha` apontando para a linha original da aba —
  pronto para colar de volta no Google Sheets.

Os ajustes ficam salvos no `localStorage` do navegador.

## Fórmulas

Com `custoLíq = custo × (1 − desconto)` e `fixos` = frete + taxas fixas em R$:

```
taxas        = NF + ADS + Devolução + Comissão

preçoMínimo  = (custoLíq + fixos) / (1 − taxas)
preçoVenda   = (custoLíq + fixos) / (1 − taxas − lucroAlvo)
preçoOriginal= preçoVenda / 0,8

lucro        = promoção − promoção × taxas − fixos − custoLíq
lucroFinal   = lucro + rebate
margem       = lucroFinal / promoção
```

O Mercado Livre tem dois canais (Clássico 11,5% e Premium 16,5%). Shopee e TikTok usam comissão
por produto, mais taxa fixa em R$ (e frete, no caso do TikTok).

## Estrutura

```
index.html          site completo, com os dados embutidos
dados/dataset.json  dados extraídos da planilha
dados/extract.py    script que gera o dataset a partir do .xlsx exportado do Google Sheets
.nojekyll           evita o processamento Jekyll no GitHub Pages
```

## Atualizar os dados

1. No Google Sheets: **Arquivo → Fazer download → Microsoft Excel (.xlsx)**.
2. Salve como `book.xlsx` ao lado do `extract.py` e rode:

   ```bash
   pip install openpyxl
   python3 dados/extract.py
   ```

3. Substitua o bloco `<script id="dataset" type="application/json">` do `index.html`
   pelo conteúdo novo de `dados/dataset.json`.

## Publicar no GitHub Pages

Em **Settings → Pages**, escolha a branch `main` e a pasta `/ (root)`.
