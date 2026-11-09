"""Gera o Comparativo F&I entre lojas (comparativo_financiamento.html) a partir das planilhas de F&I.

Uso:
    python3 gerar_comparativo_fi.py INDICADOR_DIARIO_FI.xlsx Indicador_Metas.xlsx META_FI_2026.xlsx [saida.html]

Lê só os valores já calculados (salve as planilhas no Excel antes de rodar).
Nenhum número é inventado: o que não existe nas planilhas aparece como "—" na página.

De onde vem cada coluna (por loja = coluna "estoque tipo" das bases):
- Faturamento ........ "Volume Financ. Interno (R$)" da aba RELATÓRIO do Indicador Diário
- Penetration ........ "Positivação F&I (%)" da aba RELATÓRIO do Indicador Diário
  (a aba RELATÓRIO mostra um mês por vez; para os outros meses o gerador aplica a mesma conta
   da aba às bases LINX REL VENDAS + NBS REL VENDAS: financ. interno ÷ vendas totais e soma do valor
   financiado interno, e confere que bate com a aba no mês que ela mostra)
- Principal banco .... banco com mais financiamentos INTERNOS no mês
- Retorno ............ LINX RES VEIC VEND FI "Retorno" + NBS REL VENDAS "RETORNO6"
- Retorno médio ...... retorno ÷ financiamentos internos (o "unit" da aba de resumo)
- % atingimento meta . retorno ÷ meta da loja (aba Metas do Indicador de Metas; nov = aba Painel)
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


def ler_vendas(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    base = defaultdict(lambda: defaultdict(lambda: {"vendas": 0, "financiadas": 0,
                                                    "volume": 0.0, "retorno": 0.0, "bancos": Counter()}))
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
            a["bancos"][r.get("Banco") or "SEM BANCO"] += 1

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
            a["bancos"][r.get("BANCO") or "SEM BANCO"] += 1

    for r in linhas(wb["LINX RES VEIC VEND FI"]):
        if r.get("MÊS") not in MESES or not r.get("estoque tipo"):
            continue
        reg(r)["retorno"] += num(r.get("Retorno"))

    por_mes = {}
    for mes, lojas in base.items():
        por_mes[mes] = {}
        for loja, a in lojas.items():
            por_mes[mes][loja] = {**{k: (round(v, 2) if isinstance(v, float) else v)
                                     for k, v in a.items() if k != "bancos"},
                                  "bancos": dict(a["bancos"].most_common())}
    ano = anos.most_common(1)[0][0] if anos else None
    return por_mes, ano


def ler_metas(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb["Metas"]
    cab = {str(ws.cell(4, c).value).strip().upper(): c for c in range(3, ws.max_column + 1)
           if ws.cell(4, c).value}
    metas = defaultdict(dict)
    for r in range(5, ws.max_row + 1):
        nome = ws.cell(r, 2).value
        if nome not in LOJA_META:
            continue
        for mes, c in cab.items():
            if mes in MESES and isinstance(ws.cell(r, c).value, (int, float)):
                metas[LOJA_META[nome]][mes] = ws.cell(r, c).value

    # Painel: realizado parcial do mês corrente
    p = wb["Painel"]
    titulo = str(p["A1"].value or "")
    mes_painel = next((MESES[i] for i, m in enumerate(MESES_EXTENSO)
                       if m.lower() in titulo.lower()), None)
    corte = p["C3"].value
    parcial = {
        "mes": mes_painel,
        "corte": corte.strftime("%d/%m/%Y") if hasattr(corte, "strftime") else corte,
        "diasUteis": p["C4"].value, "diasDecorridos": p["C5"].value,
        "lojas": {},
    }
    for r in range(9, p.max_row + 1):
        nome = p.cell(r, 2).value
        if nome in LOJA_META:
            parcial["lojas"][LOJA_META[nome]] = {
                "meta": p.cell(r, 3).value, "realizado": p.cell(r, 4).value,
                "projecao": p.cell(r, 6).value, "pctProjecao": p.cell(r, 7).value,
            }
    return dict(metas), parcial


def ler_meta_grupo(xlsx):
    ws = openpyxl.load_workbook(xlsx, data_only=True)["META"]
    out = []
    for r in range(2, ws.max_row + 1):
        mes = str(ws.cell(r, 2).value or "").strip().upper()
        if mes not in MESES_EXTENSO:
            continue
        val = lambda c: ws.cell(r, c).value if isinstance(ws.cell(r, c).value, (int, float)) else None
        real = val(4) or None  # célula vazia/zero = mês ainda não fechado
        out.append({"mes": MESES[MESES_EXTENSO.index(mes)], "meta": val(3), "orcado1": val(11),
                    "orcado2": val(22), "real": real})
    return out


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
                     "volume": ws.cell(r, 5).value, "positivacao": ws.cell(r, 6).value}
    return (mes if mes in MESES else None), out


def main(diario, metas_xlsx, meta_fi, saida):
    por_mes, ano = ler_vendas(diario)
    mes_rel, rel = ler_relatorio(diario)
    if mes_rel:
        for loja, v in rel.items():
            a = por_mes.setdefault(mes_rel, {}).setdefault(
                loja, {"vendas": 0, "financiadas": 0, "volume": 0, "retorno": 0, "bancos": {}})
            if (a["vendas"], a["financiadas"], round(a["volume"], 2)) != (v["vendas"], v["financiadas"], round(v["volume"] or 0, 2)):
                print(f"AVISO: {loja} {mes_rel}: bases {a['vendas']}/{a['financiadas']}/{a['volume']} "
                      f"x RELATÓRIO {v['vendas']}/{v['financiadas']}/{v['volume']} (vale o RELATÓRIO)")
            a.update(v)  # no mês da aba RELATÓRIO valem os números da própria aba
    metas, parcial = ler_metas(metas_xlsx)
    lojas = sorted({l for m in por_mes.values() for l in m} | set(metas))
    dados = {
        "ano": ano,
        "meses": [m for m in MESES if m in por_mes or m == parcial["mes"]],
        "lojas": lojas,
        "porMes": por_mes,
        "mesRelatorio": mes_rel,
        "metas": metas,
        "parcial": parcial,
        "metaGrupo": ler_meta_grupo(meta_fi),
    }
    template = (Path(__file__).parent / "template_comparativo_fi.html").read_text(encoding="utf-8")
    html = template.replace("/*__DADOS__*/null", json.dumps(dados, ensure_ascii=False))
    Path(saida).write_text(html, encoding="utf-8")
    print(f"OK: {saida}  ({len(lojas)} lojas, meses com vendas: {', '.join(m for m in MESES if m in por_mes)}; "
          f"parcial: {parcial['mes']} até {parcial['corte']})")


if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    main(*sys.argv[1:4], sys.argv[4] if len(sys.argv) > 4 else "comparativo_financiamento.html")
