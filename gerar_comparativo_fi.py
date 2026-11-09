"""Gera o Comparativo F&I entre lojas (comparativo_financiamento.html) a partir das planilhas de F&I.

Uso:
    # a partir da planilha base (o mesmo arquivo que se envia pela página em "Atualizar dados")
    python3 gerar_comparativo_fi.py BASE_COMPARATIVO_FI.xlsx [saida.html]

    # a partir das planilhas originais, em qualquer ordem (também grava BASE_COMPARATIVO_FI.xlsx preenchida)
    python3 gerar_comparativo_fi.py INDICADOR_DIARIO_FI.xlsx METAS_BI_2026.xlsx \
        CONTROLE_SOLICITACOES_NOTAS_FISCAIS.xlsx [Fichas_Aprovadas_x_Recusadas.xlsx] [saida.html]

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
- % aprovação ........ fichas aprovadas ÷ fichas passadas (Fichas Aprovadas x Recusadas, aba Dados;
                       sem testes TESTE FANDI e sem rascunhos); loja/mês sem fichas = "sem dado"
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


# ---------------------------------------------------------------------------
# Dados no formato da página (o mesmo que a planilha base guarda)
#   porMes[MES][LOJA] = {vendas, financiadas, volume, retornoRel, retorno, banco, bancoQtd, comBanco}
#   metas[LOJA][MES] = meta de retorno;  metaGrupo = [{mes, meta}];  retornoNF[MES] = real de retorno
# ---------------------------------------------------------------------------

def dados_das_planilhas(diario, metas_bi, notas, fichas=None):
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
            v = {k: x for k, x in v.items() if k != "positivacao"}  # positivação = financ. ÷ vendas
            a.update(v)  # no mês da aba RELATÓRIO valem os números da própria aba
    for lojas in por_mes.values():
        for a in lojas.values():
            bancos = a.pop("bancos")
            top = next(iter(bancos.items()), (None, 0))
            a.update(banco=top[0], bancoQtd=top[1], comBanco=sum(bancos.values()))
    for lojas in por_mes.values():
        for a in lojas.values():
            a.update(fichas=None, aprovadas=None)
    if fichas:
        for mes, lojas in ler_fichas(fichas).items():
            for loja, f in lojas.items():
                a = por_mes.setdefault(mes, {}).setdefault(loja, {
                    "vendas": 0, "financiadas": 0, "volume": 0, "retorno": 0, "retornoRel": 0,
                    "banco": None, "bancoQtd": 0, "comBanco": 0})
                a.update(f)
    metas, meta_grupo = ler_metas_bi(metas_bi)
    return {"ano": ano, "porMes": por_mes, "metas": metas, "metaGrupo": meta_grupo,
            "retornoNF": ler_retorno_nf(notas)}


def completar(dados):
    """Listas de meses e lojas a partir do conteúdo."""
    pm = dados["porMes"]
    dados["meses"] = [m for m in MESES if m in pm]
    dados["lojas"] = sorted({l for m in pm.values() for l in m} | {l for l, v in dados["metas"].items() if v})
    dados["parcial"] = None
    return dados


# ---- planilha base ---------------------------------------------------------
COLS_MES = [  # (chave, cabeçalho, de onde vem no Indicador Diário)
    ("ano", "Ano", ""),
    ("mes", "Mês", "JAN, FEV, MAR, ABR, MAI, JUN, JUL, AGO, SET, OUT, NOV ou DEZ"),
    ("loja", "Loja", "Nome da loja (igual em todas as abas)"),
    ("vendas", "Vendas totais (qtd)", "RELATÓRIO, coluna C"),
    ("financiadas", "Financ. interno (qtd)", "RELATÓRIO, coluna D"),
    ("volume", "Volume financ. interno (R$)", "RELATÓRIO, coluna E — vira o Faturamento"),
    ("retornoRel", "Retorno F&I gerado (R$)", "RELATÓRIO, coluna G — usado no Retorno médio"),
    ("retorno", "Retorno total LINX + NBS (R$)", "LINX RES VEIC VEND FI + NBS REL VENDAS — usado no % atingimento"),
    ("banco", "Principal banco", "Banco que mais aparece (LINX REL VENDAS col. D + NBS REL VENDAS col. E)"),
    ("bancoQtd", "Vendas no principal banco", "Quantas vendas da loja foram nesse banco"),
    ("comBanco", "Vendas com banco (total)", "Todas as vendas da loja com banco (sem À VISTA)"),
    ("fichas", "Fichas passadas (qtd)", "Fichas Aprovadas x Recusadas: fichas enviadas aos bancos, sem testes (TESTE FANDI) e sem rascunhos"),
    ("aprovadas", "Fichas aprovadas (qtd)", "Fichas Aprovadas x Recusadas: fichas com desfecho aprovado — vira o % aprovação"),
]
OPCIONAIS = {"fichas", "aprovadas"}  # vazio = sem dado (não vira 0)

# loja do relatório de fichas -> loja das bases de vendas
LOJA_FICHAS = {
    "MIT | MANAUS": "MITSUBISHI MANAUS", "MIT | BOA VISTA": "MITSUBISHI BOA VISTA",
    "VOLVO | MANAUS": "VOLVO MANAUS", "VOLVO | BELÉM": "VOLVO BELEM", "VOLVO | BELEM": "VOLVO BELEM",
    "MG | MANAUS": "MG MANAUS", "OMODA | BOA VISTA": "OMODA BOA VISTA", "GWM | BOA VISTA": "GWM BOA VISTA",
    "OMODA | IMPERATRIZ": "OMODA IMPERATRIZ", "GB SELECT": "GB SELECT (MANAUS, BOA VISTA E BELÉM)",
}


def ler_fichas(xlsx):
    """Fichas Aprovadas x Recusadas, aba Dados: 1 linha = 1 ficha enviada a 1 banco.
    Fora: testes (coluna "Teste?" = Sim ou cliente "TESTE FANDI") e rascunhos (não enviados)."""
    ws = openpyxl.load_workbook(xlsx, data_only=True)["Dados"]
    out = defaultdict(lambda: defaultdict(lambda: {"fichas": 0, "aprovadas": 0}))
    sem_loja = Counter()
    for r in linhas(ws):
        teste = str(r.get("Teste?") or "").strip().upper().startswith("S") or "TESTE" in str(r.get("Cliente") or "").upper()
        grupo = str(r.get("Grupo") or "").strip().upper()
        if teste or grupo == "RASCUNHO" or not grupo:
            continue
        data = r.get("Mês") or r.get("Data inclusão")
        if not hasattr(data, "month"):
            continue
        loja_txt = str(r.get("Loja") or "").strip().upper()
        loja = LOJA_FICHAS.get(loja_txt)
        if not loja:
            sem_loja[loja_txt] += 1
            continue
        a = out[MESES[data.month - 1]][loja]
        a["fichas"] += 1
        a["aprovadas"] += grupo.startswith("APROV")
    for l, n in sem_loja.items():
        print(f"AVISO: loja {l!r} das fichas sem correspondência ({n} fichas ignoradas)")
    return out


def escrever_base(dados, caminho):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    wb = openpyxl.Workbook()
    cab_fill, cab_font = PatternFill("solid", fgColor="0F2747"), Font(bold=True, color="FFFFFF")

    def cabecalho(ws, nomes, larguras):
        ws.append(nomes)
        for c, w in zip(ws[1], larguras):
            c.fill, c.font = cab_fill, cab_font
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[c.column_letter].width = w
        ws.row_dimensions[1].height = 32
        ws.freeze_panes = "A2"

    ws = wb.active
    ws.title = "LEIA-ME"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 90
    linhas_leia = [
        ("PLANILHA BASE — COMPARATIVO F&I", ""),
        ("", ""),
        ("Como atualizar", "1) Preencha as abas DADOS_MES, METAS e RETORNO_NF (uma linha por mês e loja em DADOS_MES)."),
        ("", "2) Salve o arquivo (.xlsx)."),
        ("", "3) Na página do comparativo, em “Atualizar dados”, clique em “Enviar planilha” e escolha este arquivo."),
        ("", "4) Confira a prévia e clique em “Publicar atualização”. O mesmo link passa a mostrar os dados novos."),
        ("", ""),
        ("Regras", "Não mude os nomes das abas nem dos cabeçalhos. A ordem das linhas não importa."),
        ("", "Valores em R$ sem fórmula de texto (números). Célula vazia = sem dado; meta vazia ou 0 = sem meta."),
        ("", "Um mês que você enviar substitui o que a página tinha; a planilha precisa ter todos os meses que devem aparecer."),
        ("", ""),
        ("Aba DADOS_MES — coluna", "De onde vem (Indicador Diário F&I)"),
    ] + [(nome, origem) for _, nome, origem in COLS_MES] + [
        ("", ""),
        ("Aba METAS", "Meta de retorno por loja (linhas) e mês (colunas). Linha TOTAL GRUPO = meta do grupo."),
        ("Aba RETORNO_NF", "Real de retorno do mês: Controle de Solicitações de NF, aba Resumo, quadro 4) POR TIPO, linha RETORNO."),
        ("", ""),
        ("Como a página calcula", "Penetration = Financ. interno ÷ Vendas totais. Retorno médio = Retorno F&I gerado ÷ Financ. interno."),
        ("", "% aprovação = Fichas aprovadas ÷ Fichas passadas (vazio = sem dado)."),
        ("", "% atingimento = Retorno total LINX + NBS ÷ meta da loja. Var. MA = Volume financ. interno contra o mês anterior."),
    ]
    for a, b in linhas_leia:
        ws.append([a, b])
    ws["A1"].font = Font(bold=True, size=14)
    for r in (3, 8, 12, 12 + len(COLS_MES) + 2, 12 + len(COLS_MES) + 3, 12 + len(COLS_MES) + 5):
        ws.cell(r, 1).font = Font(bold=True)

    ws = wb.create_sheet("DADOS_MES")
    cabecalho(ws, [n for _, n, _ in COLS_MES], [7, 7, 38, 11, 11, 18, 17, 18, 16, 13, 13, 12, 12])
    for mes in MESES:
        for loja in sorted(dados["porMes"].get(mes, {})):
            v = dados["porMes"][mes][loja]
            ws.append([dados["ano"], mes, loja] + [v.get(k) for k, _, _ in COLS_MES[3:]])
    for row in ws.iter_rows(min_row=2):
        for c in row[5:8]:
            c.number_format = '#,##0.00'
    dv = DataValidation(type="list", formula1='"' + ",".join(MESES) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("B2:B5000")

    ws = wb.create_sheet("METAS")
    cabecalho(ws, ["Loja"] + MESES, [38] + [11] * 12)
    for loja in sorted(set(dados["metas"]) | set(dados.get("lojas", []))):
        ws.append([loja] + [dados["metas"].get(loja, {}).get(m) for m in MESES])
    grupo = {g["mes"]: g["meta"] for g in dados["metaGrupo"]}
    ws.append(["TOTAL GRUPO"] + [grupo.get(m) for m in MESES])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2, min_col=2):
        for c in row:
            c.number_format = '#,##0'

    ws = wb.create_sheet("RETORNO_NF")
    cabecalho(ws, ["Mês", "Real de retorno (R$)"], [8, 20])
    for m in MESES:
        ws.append([m, dados["retornoNF"].get(m)])
        ws.cell(ws.max_row, 2).number_format = '#,##0.00'
    wb.active = 1
    wb.save(caminho)


def _norm(s):
    import unicodedata
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return "".join(ch for ch in s if ch.isalnum())


def _num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v
    t = str(v).strip().replace("R$", "").replace(" ", "")
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        raise ValueError(f"valor não numérico: {v!r}")


def ler_base(caminho):
    """Lê a planilha base (mesmas regras do botão "Enviar planilha" da página)."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    falta = [a for a in ("DADOS_MES", "METAS", "RETORNO_NF") if a not in wb.sheetnames]
    if falta:
        sys.exit(f"Planilha base sem as abas: {', '.join(falta)}")
    ws = wb["DADOS_MES"]
    cab = [_norm(c.value) for c in ws[1]]
    idx = {}
    for k, nome, _ in COLS_MES:
        if _norm(nome) not in cab:
            if k in OPCIONAIS:
                idx[k] = None
                continue
            sys.exit(f"DADOS_MES sem a coluna {nome!r}")
        idx[k] = cab.index(_norm(nome))
    por_mes, anos = {}, Counter()
    for n, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(v not in (None, "") for v in row):
            continue
        mes = str(row[idx["mes"]] or "").strip().upper()[:3]
        loja = str(row[idx["loja"]] or "").strip().upper()
        if mes not in MESES or not loja:
            sys.exit(f"DADOS_MES linha {n}: mês ou loja inválido")
        if row[idx["ano"]]:
            anos[int(_num(row[idx["ano"]]))] += 1
        v = {k: _num(row[idx[k]]) or 0 for k in ("vendas", "financiadas", "volume", "retornoRel", "retorno",
                                                    "bancoQtd", "comBanco")}
        v.update({k: (_num(row[idx[k]]) if idx[k] is not None else None) for k in OPCIONAIS})
        v["banco"] = str(row[idx["banco"]] or "").strip().upper() or None
        por_mes.setdefault(mes, {})[loja] = v
    metas, grupo = {}, []
    ws = wb["METAS"]
    cab = [str(c.value or "").strip().upper()[:3] for c in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        loja = str(row[0] or "").strip().upper()
        if not loja:
            continue
        vals = {m: _num(row[cab.index(m)]) for m in MESES if m in cab}
        vals = {m: v for m, v in vals.items() if v}
        if loja.startswith("TOTAL"):
            grupo = [{"mes": m, "meta": vals[m]} for m in MESES if m in vals]
        else:
            metas[loja] = vals
    if not grupo:  # sem linha TOTAL GRUPO: soma das lojas
        grupo = [{"mes": m, "meta": sum(v.get(m, 0) for v in metas.values())} for m in MESES
                 if any(m in v for v in metas.values())]
    retorno_nf = {}
    for row in wb["RETORNO_NF"].iter_rows(min_row=2, values_only=True):
        mes = str(row[0] or "").strip().upper()[:3]
        if mes in MESES and _num(row[1]):
            retorno_nf[mes] = _num(row[1])
    return {"ano": anos.most_common(1)[0][0] if anos else None, "porMes": por_mes, "metas": metas,
            "metaGrupo": grupo, "retornoNF": retorno_nf}


# ---- página ----------------------------------------------------------------
def _js(valor):
    """JSON seguro dentro de <script>."""
    return json.dumps(valor, ensure_ascii=False).replace("</", "<\\/").replace("<!--", "<\\!--")


def gerar_pagina(dados, saida):
    import re
    template = (Path(__file__).parent / "template_comparativo_fi.html").read_text(encoding="utf-8")
    # troca os dois marcadores numa passada só (a FONTE contém os marcadores e não pode ser reprocessada)
    html = re.sub(r"/\*__(DADOS|FONTE)__\*/null",
                  lambda m: _js(dados) if m.group(1) == "DADOS" else _js(template), template)
    Path(saida).write_text(html, encoding="utf-8")
    print(f"OK: {saida}  ({len(dados['lojas'])} lojas, meses: {', '.join(dados['meses'])}; "
          f"metas: {', '.join(g['mes'] for g in dados['metaGrupo'])})")


def main(args):
    xlsx = [a for a in args if a.lower().endswith((".xlsx", ".xlsm"))]
    saida = next((a for a in args if a.lower().endswith(".html")), "comparativo_financiamento.html")
    if len(xlsx) == 1:
        dados = completar(ler_base(xlsx[0]))
    elif len(xlsx) in (3, 4):
        papeis = {}
        for x in xlsx:
            abas = set(openpyxl.load_workbook(x, read_only=True).sheetnames)
            papel = ("diario" if "RELATÓRIO" in abas else "notas" if "Resumo" in abas and "Controle" in abas
                     else "fichas" if "Dados" in abas and "Mapa" in abas else "metas")
            if papel in papeis:
                sys.exit(f"Duas planilhas parecem ser {papel}: {papeis[papel]} e {x}")
            papeis[papel] = x
        if not {"diario", "metas", "notas"} <= set(papeis):
            sys.exit("Preciso do Indicador Diário, da METAS BI e do Controle de NFs (e, opcional, das Fichas).")
        dados = completar(dados_das_planilhas(papeis["diario"], papeis["metas"], papeis["notas"], papeis.get("fichas")))
        base = Path(saida).with_name("BASE_COMPARATIVO_FI.xlsx")
        escrever_base(dados, base)
        print(f"OK: {base}  (planilha base preenchida)")
    else:
        sys.exit(__doc__)
    gerar_pagina(dados, saida)


if __name__ == "__main__":
    main(sys.argv[1:])
