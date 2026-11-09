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

Página publicada: https://claude.ai/artifact/VdxkL4QVbPLk9HwsPmrkZY

`comparativo_financiamento.html` compara as lojas mês a mês: faturamento (volume financ. interno), Var. MA,
% aprovação, penetration, principal banco, % atingimento da meta de retorno e retorno médio, mais a meta
de retorno do grupo contra o real de retorno das notas fiscais. Tem botão **Gerar PDF** (A4 paisagem).

### Atualizar pela própria página (sem rodar nada)

1. Em **Atualizar dados**, clique em **Baixar planilha base** (vem com os dados atuais).
2. Atualize as abas `DADOS_MES`, `METAS` e `RETORNO_NF` no Excel e salve.
3. Clique em **Enviar planilha atualizada**: a tabela mostra a prévia e a página aponta erros.
4. Clique em **Publicar atualização**: o mesmo link passa a mostrar os dados novos para todos
   (só quem tem permissão de edição na página consegue publicar).

A aba LEIA-ME da planilha base diz de onde vem cada coluna (Indicador Diário F&I, METAS BI 2026 e
Controle de Solicitações de NF).

### Atualizar pelo gerador

```bash
# a partir da planilha base
python3 gerar_comparativo_fi.py BASE_COMPARATIVO_FI.xlsx comparativo_financiamento.html

# a partir das planilhas originais (também grava BASE_COMPARATIVO_FI.xlsx preenchida)
python3 gerar_comparativo_fi.py INDICADOR_DIARIO_FI_GB_NORTE.xlsx METAS_BI_2026.xlsx \
    CONTROLE_SOLICITACOES_NOTAS_FISCAIS.xlsx Fichas_Aprovadas_x_Recusadas_por_Loja.xlsx comparativo_financiamento.html
```

As planilhas originais podem vir em qualquer ordem (o gerador reconhece cada uma pelas abas).
**% aprovação** = fichas aprovadas ÷ fichas passadas (aba Dados das Fichas, sem testes TESTE FANDI e sem rascunhos).

O arquivo gerado já está no formato de publicação (sem `<html>`/`<head>`): ele leva uma cópia do
`template_comparativo_fi.html` para conseguir se republicar com dados novos.

## Arquivos

- `gerar_one_page.py` — lê a planilha e grava o HTML.
- `template.html` — layout da página (os dados entram no lugar de `/*__DADOS__*/null`).
- `index.html` — página gerada.
- `gerar_comparativo_fi.py` + `template_comparativo_fi.html` — geram `comparativo_financiamento.html` (comparativo de F&I entre lojas) e a planilha base `BASE_COMPARATIVO_FI.xlsx`.
