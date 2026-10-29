"""Gera a ONE PAGE de Pós-Vendas (index.html) a partir da planilha mãe.

Uso:
    python3 gerar_one_page.py INDICADORES_POS_VENDA_planilha_mae.xlsx [saida.html]

Lê os valores já calculados pelo Excel (salve a planilha no Excel antes de rodar)
e embute os dados em template.html.
"""
import json
import sys
from pathlib import Path

import openpyxl

MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
MESES_EXTENSO = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
                 "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

# aba -> (chave, rótulo, unidade, como consolidar)
INDICADORES = [
    ("FATURAMENTO", "fat", "Faturamento", "R$", "soma"),
    ("MARGEM BRUTA", "margem", "Margem bruta", "%", "media_pond_fat"),
    ("PASSAGENS", "pass", "Passagens", "qtd", "soma"),
    ("TICKET MÉDIO", "ticket", "Ticket médio", "R$", "media_pond_pass"),
    ("NPS", "nps", "NPS", "pts", "media"),
    ("FFV", "ffv", "FFV", "%", "media"),
    ("DID", "did", "DID", "%", "media"),
]


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def bloco(ws, linha_ini, n_lojas):
    """Lê lojas (col B) e 12 meses (cols C:N) a partir de linha_ini; última linha = consolidado."""
    lojas, cons = {}, None
    for r in range(linha_ini, linha_ini + n_lojas + 1):
        nome = ws.cell(r, 2).value
        vals = [num(ws.cell(r, c).value) for c in range(3, 15)]
        if r == linha_ini + n_lojas:
            cons = vals
        else:
            lojas[nome] = vals
    return lojas, cons


def main(xlsx, saida):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    cfg = wb["CONFIGURAÇÃO"]
    lojas = [cfg.cell(r, 3).value for r in range(12, 21) if cfg.cell(r, 3).value]
    n = len(lojas)

    fontes, faixas = {}, {}
    for r in range(1, cfg.max_row + 1):
        rot = str(cfg.cell(r, 2).value or "")
        if 30 <= r <= 40 and cfg.cell(r, 3).value and not rot.startswith(("Verde", "Amarelo")):
            fontes[rot] = cfg.cell(r, 3).value
        if rot.startswith("Verde"):
            faixas["verde"] = cfg.cell(r, 3).value
        if rot.startswith("Amarelo"):
            faixas["amarela"] = cfg.cell(r, 3).value

    dados = {
        "empresa": cfg["C6"].value,
        "ano": cfg["C7"].value,
        "periodo1": cfg["C23"].value,
        "periodo2": cfg["C24"].value,
        "inicioPeriodo2": MESES.index(cfg["C25"].value) if cfg["C25"].value in MESES else None,
        "faixaVerde": faixas.get("verde") or 1,
        "faixaAmarela": faixas.get("amarela") or 0.9,
        "lojas": lojas,
        "meses": MESES,
        "indicadores": [],
    }

    for aba, chave, rotulo, unidade, consol in INDICADORES:
        if aba not in wb.sheetnames:
            continue
        ws = wb[aba]
        real, real_cons = bloco(ws, 8, n)
        meta, meta_cons = bloco(ws, 21, n)
        comentarios = {ws.cell(r, 2).value: ws.cell(r, 3).value
                       for r in range(47, 47 + n + 1) if ws.cell(r, 3).value}
        dados["indicadores"].append({
            "chave": chave, "rotulo": rotulo, "unidade": unidade, "consolidacao": consol,
            "fonte": fontes.get(rotulo, ""),
            "real": real, "realCons": real_cons, "meta": meta, "metaCons": meta_cons,
            "comentarios": comentarios,
        })

    # Oficina: exports Linx/NBS de cada mês; a loja de cada linha vem da aba LIGAÇÃO EXPORTS
    # (mesma linha do export). Colunas: chave -> (col Linx, col NBS)
    campos = {"passagens": (2, 2), "mo": (3, 3), "objMo": (4, None), "pecas": (7, 4),
              "objPecas": (8, None), "total": (11, 5), "projecao": (13, 7), "lbPecas": (15, 9)}
    lig = wb["LIGAÇÃO EXPORTS"]
    linhas = []
    for aba, origem, col_mes, col_nome, col_loja, idx in (
            ("PASSAGENS LINXS", "Linx", 1, 2, 3, 0), ("PASSAGENS NBS", "NBS", 7, 8, 9, 1)):
        ws = wb[aba]
        for r in range(5, lig.max_row + 1):
            mes, nome, loja = lig.cell(r, col_mes).value, lig.cell(r, col_nome).value, lig.cell(r, col_loja).value
            if not isinstance(mes, (int, float)) or not nome or ws.cell(r, 1).value != nome:
                continue  # faixa do mês ou linha vazia
            reg = {k: (num(ws.cell(r, c[idx]).value) if c[idx] else None) for k, c in campos.items()}
            if all(v is None for v in reg.values()):
                continue  # faixa "MÊS — cole aqui..." ou linha sem números
            linhas.append({"mes": int(mes) - 1, "origem": origem, "nome": nome,
                           "loja": loja or "SEM LOJA", **reg})

    def resumo(regs):
        t = {k: sum(x[k] or 0 for x in regs) for k in campos}
        for k in ("objMo", "objPecas"):
            if not any(x[k] for x in regs):
                t[k] = None
        div = lambda a, b: a / b if a is not None and b else None
        t.update(pctMo=div(t["mo"], t["objMo"]), pctPecas=div(t["pecas"], t["objPecas"]),
                 ticket=div(t["total"], t["passagens"]), margemPecas=div(t["lbPecas"], t["pecas"]),
                 pecasMo=div(t["pecas"], t["mo"]), colaboradores=len(regs))
        return t

    oficina, ranking = {}, {}
    for m in sorted({x["mes"] for x in linhas}):
        do_mes = [x for x in linhas if x["mes"] == m]
        nas_lojas = [x for x in do_mes if x["loja"] in lojas]
        oficina[m] = {
            "porLoja": {l: resumo([x for x in do_mes if x["loja"] == l]) for l in lojas
                        if any(x["loja"] == l for x in do_mes)},
            "cons": resumo(nas_lojas),
            "semLoja": sorted({x["nome"] for x in do_mes if x["loja"] not in lojas}),
        }
        rk = [{"origem": x["origem"], "nome": x["nome"], "loja": x["loja"], "passagens": x["passagens"],
               "total": x["total"], "ticket": x["total"] / x["passagens"] if x["passagens"] else None}
              for x in nas_lojas if x["total"]]
        ranking[m] = sorted(rk, key=lambda x: -x["total"])
    dados["oficina"] = {"meses": oficina, "ranking": ranking}

    # Estoque de peças (aba compra de peças: B=mês, D:G saldos, L=loja one page)
    cp = wb["compra de peças"]
    estoque = []
    for r in range(8, cp.max_row + 1):
        mes, loja = cp.cell(r, 2).value, cp.cell(r, 12).value
        if not mes or not loja or str(mes).strip().upper() not in MESES_EXTENSO:
            continue
        if not any(cp.cell(r, c).value for c in range(4, 8)):
            continue  # linha do mês ainda zerada (sem lançamento)
        estoque.append({
            "mes": MESES_EXTENSO.index(str(mes).strip().upper()), "loja": loja,
            "inicial": num(cp.cell(r, 4).value), "entradas": num(cp.cell(r, 5).value),
            "saidas": num(cp.cell(r, 6).value), "final": num(cp.cell(r, 7).value),
        })
    dados["estoque"] = estoque

    # Qualidade montadora
    qm = wb["QUALIDADE MONTADORA"]
    qualidade = []
    for r in range(7, qm.max_row + 1):
        if not qm.cell(r, 3).value:
            continue
        qualidade.append({
            "marca": qm.cell(r, 2).value, "loja": qm.cell(r, 3).value,
            "programa": qm.cell(r, 4).value, "meta": num(qm.cell(r, 5).value),
            "resultado": num(qm.cell(r, 6).value), "status": qm.cell(r, 7).value,
            "obs": qm.cell(r, 8).value,
        })
    dados["qualidade"] = {"periodo": qm["C4"].value, "itens": qualidade}

    template = (Path(__file__).parent / "template.html").read_text(encoding="utf-8")
    html = template.replace("/*__DADOS__*/null", json.dumps(dados, ensure_ascii=False))
    Path(saida).write_text(html, encoding="utf-8")
    print(f"OK: {saida}  ({len(lojas)} lojas, {len(linhas)} linhas de export em {len(oficina)} meses, {len(estoque)} linhas de estoque)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "index.html")
