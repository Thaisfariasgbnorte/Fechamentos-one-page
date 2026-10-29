# One Page Pós-Vendas — Grupo Bastazini

Página única (HTML) com os indicadores de pós-vendas gerada a partir da planilha mãe
`INDICADORES_POS_VENDA_planilha_mae.xlsx`.

## Como atualizar

1. Faça os lançamentos do mês na planilha (veja a aba LEIA-ME) e **salve no Excel**
   (o gerador lê os valores já calculados pelas fórmulas).
2. Rode:

   ```bash
   pip install openpyxl
   python3 gerar_one_page.py caminho/da/planilha.xlsx index.html
   ```

3. Abra `index.html` no navegador.

## O que a página mostra

- **Indicadores do mês**: faturamento, margem bruta, passagens, ticket médio, NPS, FFV e DID —
  realizado, meta, % da meta (verde/amarelo/vermelho pelas faixas da CONFIGURAÇÃO),
  variação vs mês anterior e acumulado/média do ano.
- **1/2 Resultados**: desempenho por loja, evolução mensal (meta × realizado) por indicador,
  conclusões automáticas e comentários lançados nas abas.
- **2/2 Oficina**: mão de obra × peças por loja (export Linx/NBS), estoque de peças,
  top 10 consultores e programas de qualidade da montadora.

Escolha o mês e a visão (consolidado ou uma loja) no topo; clicar numa loja da tabela também muda a visão.

## Arquivos

- `gerar_one_page.py` — lê a planilha e grava o HTML.
- `template.html` — layout da página (os dados entram no lugar de `/*__DADOS__*/null`).
- `index.html` — página gerada.
