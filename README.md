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

- **Indicadores do mês**: faturamento, margem bruta, passagens, ticket médio, FFV e DID (e NPS, se a aba existir) —
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

**Exportar PDF** gera 2 páginas A4 paisagem (resultados; oficina, estoque e qualidade) do período e da
visão escolhidos. Escolhendo uma loja, todas as tabelas mostram só aquela loja.

## Comparativo entre lojas (F&I)

`comparativo_financiamento.html` compara as lojas mês a mês em faturamento, % aprovação, penetration,
principal banco, % atingimento da meta e retorno médio, só com números das planilhas de F&I:

```bash
python3 gerar_comparativo_fi.py INDICADOR_DIARIO_FI_GB_NORTE.xlsx Indicador_Metas_Novembro_2026.xlsx \
    META_FI_2026.xlsx CONTROLE_SOLICITACOES_NOTAS_FISCAIS.xlsx comparativo_financiamento.html
```

- **Faturamento**: "Volume Financ. Interno (R$)" da aba RELATÓRIO do Indicador Diário.
- **Penetration**: "Positivação F&I (%)" da aba RELATÓRIO do Indicador Diário.
  A aba mostra um mês por vez; nos outros meses vale a mesma conta aplicada às bases LINX e NBS.
- **Principal banco**: banco que mais aparece em LINX REL VENDAS col. D e NBS REL VENDAS col. E
  (loja = col. K / col. M "estoque tipo"); ignora À VISTA, vazio e EXTERNO.
- **Retorno médio**: "Retorno F&I Gerado (R$)" (col. G) ÷ "Vendas Totais (Qtd)" (col. C) da aba RELATÓRIO;
  nos outros meses, a mesma conta com o retorno da LINX RES VEIC VEND FI.
- **% atingimento meta**: retorno ÷ meta da loja (aba Metas: out, nov, dez; novembro parcial pela aba Painel).
- **% aprovação**: nenhuma planilha tem fichas enviadas/aprovadas, então a coluna fica "sem dado".
- Embaixo, a meta F&I do grupo (META F&I 2026: meta e realizado) e o real de retorno
  (linha RETORNO do quadro "4) POR TIPO" da aba Resumo do controle de notas fiscais).

## Arquivos

- `gerar_one_page.py` — lê a planilha e grava o HTML.
- `template.html` — layout da página (os dados entram no lugar de `/*__DADOS__*/null`).
- `index.html` — página gerada.
- `gerar_comparativo_fi.py` + `template_comparativo_fi.html` — geram `comparativo_financiamento.html` (comparativo de F&I entre lojas).
