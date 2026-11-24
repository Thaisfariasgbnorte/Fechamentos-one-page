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

- **Indicadores do mês**: faturamento, margem bruta, passagens, ticket médio, Retorno e Chamado (abas RETORNO e CHAMADO; antes FFV e DID) —
  realizado, meta, % da meta (verde/amarelo/vermelho pelas faixas da CONFIGURAÇÃO),
  variação vs mês anterior e acumulado/média do ano.
- **1/2 Resultados**: desempenho por loja, evolução mensal (meta × realizado) por indicador,
  conclusões automáticas e comentários lançados nas abas.
- **2/2 Oficina**: mão de obra × peças por loja e top 10 consultores do mês escolhido
  (exports Linx/NBS de cada mês + loja da aba LIGAÇÃO EXPORTS), estoque de peças e
  programas de qualidade da montadora. Mês sem export ou estoque usa o último lançado.

No topo, escolha o período e a visão (consolidado ou uma loja); clicar numa loja da tabela também muda a visão.
O período pode ser um atalho (último mês, 1º a 4º trimestre, ano todo) ou qualquer combinação de meses:
clicar num mês soma ou tira esse mês. Com vários meses, faturamento e passagens são somados, a margem é
ponderada pelo faturamento, o ticket pelas passagens e os demais indicadores usam média simples; a
variação compara com o período anterior de mesmo tamanho (em p.p. para indicadores em %).

**Ordenar**: clique no título de qualquer coluna das tabelas para ordenar do maior para o menor; clique de
novo para menor→maior e uma terceira vez para voltar à ordem original. No ranking de consultores, os títulos
Pass., Total oficina e Ticket escolhem por qual valor o top 10 é montado (maiores ou menores).

**Exportar PDF** gera 2 páginas A4 paisagem (resultados; oficina, estoque e qualidade) do período e da
visão escolhidos. Escolhendo uma loja, todas as tabelas mostram só aquela loja.

**Comentários e ações do gerente**: único campo editável da página, um texto por período e visão.
No claude.ai fica salvo para todos que abrem a página; aberto como arquivo, só no próprio navegador.

**2026 × 2025**: a aba BASE 2025 (depois de TICKET MÉDIO) tem, para faturamento, margem, passagens e
ticket, um quadro amarelo para digitar 2025 e o quadro "VARIAÇÃO 2026 x 2025" calculado mês a mês.
Com 2025 preenchido, os cartões mostram a variação "vs 2025" e o gráfico mensal ganha a linha tracejada de 2025.

Porcentagens aparecem sem casa decimal (98,8% → 99%).

**Nomes de consultores**: aparecem num formato só ("Karina de Carvalho Pereira"). Grafias diferentes da
mesma pessoa se juntam pelo quadro CONSULTORES da aba DE-PARA (colunas I:J: nome no export → nome padrão).

## Arquivos

- `gerar_one_page.py` — lê a planilha e grava o HTML.
- `template.html` — layout da página (os dados entram no lugar de `/*__DADOS__*/null`).
- `index.html` — página gerada.
