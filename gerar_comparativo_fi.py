"""Gera o Comparativo F&I entre lojas (comparativo_financiamento.html) a partir das planilhas de F&I.

Uso:
    python3 gerar_comparativo_fi.py INDICADOR_DIARIO_FI.xlsx METAS_BI_2026.xlsx \
        CONTROLE_SOLICITACOES_NOTAS_FISCAIS.xlsx [saida.html]

Lê só os valores já calculados (salve as planilhas no Excel antes de rodar).
Nenhum número é inventado: o que não existe nas planilhas aparece como "—" na página.

De onde vem cada coluna (por loja = coluna "estoque tipo" das bases):
- Faturamento ........ "Volume Financ. Interno (R$)" da aba RELATÓRIO do Indicador Diário
- Penetration ........ "Positivação F&I (%)" da aba RELATÓRIO do Indicador Diário
  (a aba RELATÓRIO mostra um mês por vez; para os outros meses o gerador aplica a mesma conta
   da aba às bases LINX REL VENDAS + NBS REL VENDAS: financ. interno ÷ vendas totais e soma do valor
   financiado interno, e confere que bate com a aba no mês que ela mostra)
- Principal banco .... banco que mais aparece no mês: LINX REL VENDAS col D (Banco) e NBS REL VENDAS
                       col E (BANCO), loja = col K / col M ("estoque tipo"); ignora À VISTA, vazio e EXTERNO
- Retorno ............ LINX RES VEIC VEND FI "Retorno" + NBS REL VENDAS "RETORNO6"
- Retorno médio ...... "Retorno F&I Gerado (R$)" (col G) ÷ "Financ. Interno (Qtd)" (col D) da aba RELATÓRIO;
                       nos outros meses, a mesma conta com o retorno da LINX RES VEIC VEND FI (é o que a col G soma)
- % atingimento meta . retorno ÷ meta da loja (METAS BI 2026; meta 0 = sem meta)
- Meta do grupo ...... linha de total da METAS BI 2026 x real de retorno (aba Resumo do controle de NFs)
- % aprovação ........ as planilhas não têm fichas enviadas/aprovadas -> fica "—"
"""
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

MESES = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
MESES_EXTENSO = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
                 "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

# nome na planilha de metas -> nome da loja nas bases de vendas
LOJA_META = {
    "MIT MANAUS": "MITSUBISHI MANAUS",
    "VOLVO MANAUS": "VOLVO MANAUS",
    "MG MANAUS": "MG MANAUS",
    "VOLVO BELÉM": "VOLVO BELEM",
    "VOLVO BELEM": "VOLVO BELEM",
    "GB SELECT": "GB SELECT (MANAUS, BOA VISTA E BELÉM)",
    "MIT BOA VISTA": "MITSUBISHI BOA VISTA",
    "OMODA BOA VISTA": "OMODA BOA VISTA",
    "GWM BOA VISTA": "GWM BOA VISTA",
    "TOCAUTO": "VW TOCAUTO AUTO",
    "MAN IMPERATRIZ": "VW CAMINHÕES IMPERATRIZ",
    "MAN BALSAS": "VW CAMINHÕES BALSAS",
    "OMODA IMPERATRIZ": "OMODA IMPERATRIZ",
}


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


def linhas(ws):
    it = ws.iter_rows(values_only=True)
    cab = next(it)
    return [dict(zip(cab, r)) for r in it if any(v is not None for v in r)]


SEM_BANCO = {"", "À VISTA", "A VISTA", "EXTERNO"}


def conta_banco(a, banco):
    """Principal banco: toda venda com banco na coluna de banco, qualquer que seja o tipo."""
    banco = str(banco or "").strip().upper()
    if banco not in SEM_BANCO:
        a["bancos"][banco] += 1


def ler_vendas(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    base = defaultdict(lambda: defaultdict(lambda: {"vendas": 0, "financiadas": 0,
                                                    "volume": 0.0, "retorno": 0.0, "retornoRel": 0.0,
                                                    "bancos": Counter()}))
    anos = Counter()

    def reg(r, loja_col="estoque tipo"):
        return base[r["MÊS"]][r[loja_col]]

    for r in linhas(wb["LINX REL VENDAS"]):
        if r.get("MÊS") not in MESES or not r.get("estoque tipo"):
            continue
        anos[r.get("ANO")] += 1
        a = reg(r)
        a["vendas"] += 1
        if r.get("Tipo") == "INTERNO":
            a["financiadas"] += 1
            a["volume"] += num(r.get("Valor Financ"))
        conta_banco(a, r.get("Banco"))  # col D (loja = col K "estoque tipo")

    for r in linhas(wb["NBS REL VENDAS"]):
        if r.get("MÊS") not in MESES or not r.get("estoque tipo"):
            continue
        anos[r.get("ANO")] += 1
        a = reg(r)
        a["vendas"] += 1
        a["retorno"] += num(r.get("RETORNO6"))
        if r.get("Tipo") == "INTERNO":
            a["financiadas"] += 1
            a["volume"] += num(r.get("VAL. FINANC"))
        conta_banco(a, r.get("BANCO"))  # col E (loja = col M "estoque tipo")

    for r in linhas(wb["LINX RES VEIC VEND FI"]):
        if r.get("MÊS") not in MESES or not r.get("estoque tipo"):
            continue
        reg(r)["retorno"] += num(r.get("Retorno"))
        reg(r)["retornoRel"] += num(r.get("Retorno"))  # mesma base da col G da aba RELATÓRIO

    por_mes = {}
    for mes, lojas in base.items():
        por_mes[mes] = {}
        for loja, a in lojas.items():
            por_mes[mes][loja] = {**{k: (round(v, 2) if isinstance(v, float) else v)
                                     for k, v in a.items() if k != "bancos"},
                                  "bancos": dict(a["bancos"].most_common())}
    ano = anos.most_common(1)[0][0] if anos else None
    return por_mes, ano


def ler_metas_bi(xlsx):
    """METAS BI 2026: linha 1 = meses (SET, OUT...), coluna C = loja; linha sem loja = total do grupo.
    Meta 0 conta como "sem meta" (não dá para medir atingimento)."""
    ws = openpyxl.load_workbook(xlsx, data_only=True).worksheets[0]
    cab = {c: str(ws.cell(1, c).value).strip().upper() for c in range(1, ws.max_column + 1)
           if str(ws.cell(1, c).value or "").strip().upper() in MESES}
    metas, grupo = defaultdict(dict), {}
    for r in range(2, ws.max_row + 1):
        nome = str(ws.cell(r, 3).value or "").strip().upper()
        vals = {mes: ws.cell(r, c).value for c, mes in cab.items() if isinstance(ws.cell(r, c).value, (int, float))}
        if not nome:
            grupo.update(vals)
        elif nome in LOJA_META:
            metas[LOJA_META[nome]].update({m: v for m, v in vals.items() if v})
        elif vals:
            print(f"AVISO: loja {nome!r} da METAS BI sem correspondência nas bases (ignorada)")
    meta_grupo = [{"mes": m, "meta": grupo[m]} for m in MESES if m in grupo]
    return dict(metas), meta_grupo


def ler_retorno_nf(xlsx):
    """Linha RETORNO do quadro "4) POR TIPO" da aba Resumo do controle de notas fiscais (meses 1..12)."""
    ws = openpyxl.load_workbook(xlsx, data_only=True)["Resumo"]
    quadro = None
    for r in range(1, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        if a[:2].rstrip(")").isdigit() and ")" in a[:3]:
            quadro = a
        if quadro and quadro.startswith("4)") and a.upper() == "RETORNO":
            return {MESES[c - 2]: ws.cell(r, c).value for c in range(2, 14)
                    if isinstance(ws.cell(r, c).value, (int, float)) and ws.cell(r, c).value}
    sys.exit("Não achei a linha RETORNO do quadro 4) POR TIPO na aba Resumo do controle de notas fiscais.")


def ler_relatorio(xlsx):
    """Bloco do topo da aba RELATÓRIO: mês em L3, uma linha por loja até TOTAL GERAL."""
    ws = openpyxl.load_workbook(xlsx, data_only=True)["RELATÓRIO"]
    mes = str(ws["L3"].value or "").strip().upper()
    out = {}
    for r in range(3, ws.max_row + 1):
        loja = ws.cell(r, 2).value
        if not loja or str(loja).upper().startswith("TOTAL"):
            break
        out[loja] = {"vendas": ws.cell(r, 3).value, "financiadas": ws.cell(r, 4).value,
                     "volume": ws.cell(r, 5).value, "positivacao": ws.cell(r, 6).value,
                     "retornoRel": num(ws.cell(r, 7).value)}
    return (mes if mes in MESES else None), out


def main(diario, metas_bi, notas, saida):
    por_mes, ano = ler_vendas(diario)
    mes_rel, rel = ler_relatorio(diario)
    if mes_rel:
        for loja, v in rel.items():
            a = por_mes.setdefault(mes_rel, {}).setdefault(
                loja, {"vendas": 0, "financiadas": 0, "volume": 0, "retorno": 0, "retornoRel": 0, "bancos": {}})
            if ((a["vendas"], a["financiadas"], round(a["volume"], 2), round(a["retornoRel"], 2))
                    != (v["vendas"], v["financiadas"], round(v["volume"] or 0, 2), round(v["retornoRel"], 2))):
                print(f"AVISO: {loja} {mes_rel}: bases {a['vendas']}/{a['financiadas']}/{a['volume']} "
                      f"/{a['retornoRel']} x RELATÓRIO {v['vendas']}/{v['financiadas']}/{v['volume']}/{v['retornoRel']} (vale o RELATÓRIO)")
            a.update(v)  # no mês da aba RELATÓRIO valem os números da própria aba
    metas, meta_grupo = ler_metas_bi(metas_bi)
    lojas = sorted({l for m in por_mes.values() for l in m} | set(metas))
    dados = {
        "ano": ano,
        "meses": [m for m in MESES if m in por_mes],
        "lojas": lojas,
        "porMes": por_mes,
        "mesRelatorio": mes_rel,
        "metas": metas,
        "parcial": None,
        "metaGrupo": meta_grupo,
        "retornoNF": ler_retorno_nf(notas),
    }
    template = (Path(__file__).parent / "template_comparativo_fi.html").read_text(encoding="utf-8")
    html = template.replace("/*__DADOS__*/null", json.dumps(dados, ensure_ascii=False))
    Path(saida).write_text(html, encoding="utf-8")
    print(f"OK: {saida}  ({len(lojas)} lojas, meses com vendas: {', '.join(m for m in MESES if m in por_mes)}; "
          f"metas: {', '.join(g['mes'] for g in meta_grupo)})")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    main(*sys.argv[1:4], sys.argv[4] if len(sys.argv) > 4 else "comparativo_financiamento.html")
