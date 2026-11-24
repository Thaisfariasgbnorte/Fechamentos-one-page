"""Gera a ONE PAGE de Pós-Vendas (index.html) a partir da planilha mãe.

Uso:
    python3 gerar_one_page.py INDICADORES_POS_VENDA_planilha_mae.xlsx [saida.html]

Lê os valores já calculados pelo Excel (salve a planilha no Excel antes de rodar)
e embute os dados em template.html.
"""
import json
import sys
import unicodedata
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
    (("RETORNO", "FFV"), "ffv", "Retorno", "%", "media"),
    (("CHAMADO", "DID"), "did", "Chamado", "%", "media"),
]


# grafias diferentes da mesma pessoa (chave sem acento, maiúscula) -> nome padrão;
# preenchido pelo quadro CONSULTORES da aba DE-PARA (colunas I:J)
NOMES_IGUAIS = {}
MINUSCULAS = {"de", "da", "do", "das", "dos", "e"}
_padrao = {}


def nome_consultor(nome):
    """Nome do consultor num formato só ("Karina de Carvalho Pereira"); variações de maiúsculas,
    acentos, espaços e as grafias de NOMES_IGUAIS viram o mesmo nome."""
    bruto = " ".join(str(nome or "").split())
    if not bruto:
        return "—"
    chave = unicodedata.normalize("NFD", bruto).encode("ascii", "ignore").decode().upper()
    chave = NOMES_IGUAIS.get(chave, chave)
    if chave not in _padrao:  # a primeira grafia vista (com acento, se tiver) define o nome exibido
        base = bruto if NOMES_IGUAIS.get(chave) is None and chave == unicodedata.normalize(
            "NFD", bruto).encode("ascii", "ignore").decode().upper() else chave
        _padrao[chave] = " ".join(p.lower() if i and p.lower() in MINUSCULAS else p.capitalize()
                                  for i, p in enumerate(base.lower().split()))
    return _padrao[chave]


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
    # a CONFIGURAÇÃO é lida pelos rótulos da coluna B, então lojas novas não quebram o gerador
    lojas, fontes, faixas, cfg_val = [], {}, {}, {}
    em_fontes = False
    for r in range(1, cfg.max_row + 1):
        rot = str(cfg.cell(r, 2).value or "").strip()
        val = cfg.cell(r, 3).value
        if rot.startswith("Loja ") and val:
            lojas.append(val)
        elif rot == "Indicador":
            em_fontes = True
            continue
        elif rot.startswith("FAIXAS"):
            em_fontes = False
        if em_fontes and rot and val:
            fontes[rot] = val
        if rot.startswith("Verde"):
            faixas["verde"] = val
        if rot.startswith("Amarelo"):
            faixas["amarela"] = val
        cfg_val[rot] = val
    n = len(lojas)

    # de-para de consultores (DE-PARA, colunas I:J): várias grafias -> um nome padrão
    if "DE-PARA" in wb.sheetnames:
        dp = wb["DE-PARA"]
        if "CONSULTORES" in str(dp["I1"].value or ""):
            chave = lambda t: unicodedata.normalize("NFD", " ".join(str(t).split())).encode("ascii", "ignore").decode().upper()
            for r in range(2, dp.max_row + 1):
                de, para = dp.cell(r, 9).value, dp.cell(r, 10).value
                if de and para:
                    NOMES_IGUAIS[chave(de)] = chave(para)

    dados = {
        "empresa": cfg["C6"].value,
        "ano": cfg["C7"].value,
        "periodo1": cfg_val.get("Nome do período 1"),
        "periodo2": cfg_val.get("Nome do período 2"),
        "inicioPeriodo2": MESES.index(cfg_val.get("Mês em que começa o período 2"))
                          if cfg_val.get("Mês em que começa o período 2") in MESES else None,
        "faixaVerde": faixas.get("verde") or 1,
        "faixaAmarela": faixas.get("amarela") or 0.9,
        "lojas": lojas,
        "meses": MESES,
        "indicadores": [],
    }

    for aba, chave, rotulo, unidade, consol in INDICADORES:
        # aba pode ter mais de um nome (ex.: RETORNO, antes FFV)
        aba = next((a for a in ((aba,) if isinstance(aba, str) else aba) if a in wb.sheetnames), None)
        if not aba:
            continue
        ws = wb[aba]
        # seções localizadas pelo rótulo da coluna B (META, COMENTÁRIOS); dados 2 linhas abaixo
        secao = {str(ws.cell(r, 2).value or "").split(" ")[0]: r for r in range(1, ws.max_row + 1)}
        real, real_cons = bloco(ws, 8, n)
        meta, meta_cons = bloco(ws, secao["META"] + 2, n)
        ini_com = secao.get("COMENTÁRIOS", ws.max_row) + 2
        comentarios = {ws.cell(r, 2).value: ws.cell(r, 3).value
                       for r in range(ini_com, ini_com + n + 1) if ws.cell(r, 3).value}
        # realizado 2025 (base da comparação 2026 x 2025), quando a planilha tiver
        real25, real25_cons = {}, [None] * 12
        # aba BASE 2025: um quadro por indicador, com título "<ABA> 2025 ..."
        if "BASE 2025" in wb.sheetnames:
            b25 = wb["BASE 2025"]
            lin25 = next((r for r in range(1, b25.max_row + 1)
                          if str(b25.cell(r, 2).value or "").startswith(f"{aba} 2025")), None)
            if lin25:
                real25, real25_cons = bloco(b25, lin25 + 2, n)
                # 0 em 2025 = loja ainda não aberta: fica sem comparação naquele mês
                real25 = {l: [None if v == 0 else v for v in vs] for l, vs in real25.items()}
                real25_cons = [None if v == 0 else v for v in real25_cons]
        dados["indicadores"].append({
            "chave": chave, "rotulo": rotulo, "unidade": unidade, "consolidacao": consol,
            "fonte": fontes.get(rotulo, fontes.get(rotulo.upper(), "")),
            "real": real, "realCons": real_cons, "meta": meta, "metaCons": meta_cons,
            "comentarios": comentarios,
            "real2025": real25, "real2025Cons": real25_cons,
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
            if str(nome).strip().upper() in ("TOTAL", "NOME") or str(nome).startswith("Filtros "):
                continue  # linha de total/cabeçalho do export (a LIGAÇÃO EXPORTS também ignora)
            linhas.append({"mes": int(mes) - 1, "origem": origem, "nome": nome_consultor(nome),
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

    # OS negadas: uma aba "OSN ..." por loja (o relatório não traz a loja); a loja vem do título A1
    def valor_br(v):
        if isinstance(v, (int, float)):
            return v
        t = str(v or "").strip().replace("R$", "").replace(" ", "")
        if not t:
            return None
        if "," in t and "." in t:   # o separador que vem por último é o dos centavos
            t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") else t.replace(",", "")
        elif "," in t:
            t = t.replace(",", "") if len(t.split(",")[-1]) == 3 else t.replace(",", ".")
        try:
            return float(t)
        except ValueError:
            return None
    os_negada = {}
    for aba in wb.sheetnames:
        if not aba.startswith("OSN "):
            continue
        ws = wb[aba]
        loja = str(ws["A1"].value or "").split("—")[-1].strip()
        if loja not in lojas:
            continue
        meses = {}
        for r in range(5, ws.max_row + 1):
            data = ws.cell(r, 3).value
            if not hasattr(data, "year") or data.year != dados["ano"]:
                continue
            m = meses.setdefault(data.month - 1, {"qtd": 0, "valor": 0.0, "vendedores": {}})
            v = valor_br(ws.cell(r, 8).value) or 0.0
            m["qtd"] += 1
            m["valor"] += v
            vend = nome_consultor(ws.cell(r, 7).value)
            q = m["vendedores"].setdefault(vend, [0, 0.0])
            q[0] += 1
            q[1] += v
        if meses:  # aba vazia = relatório ainda não lançado
            os_negada[loja] = meses
    dados["osNegada"] = os_negada

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
    periodo = qm["C4"].value
    if hasattr(periodo, "month"):  # data: vira "outubro/2026"
        nomes = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
                 "setembro", "outubro", "novembro", "dezembro"]
        periodo = f"{nomes[periodo.month - 1]}/{periodo.year}"
    dados["qualidade"] = {"periodo": periodo, "itens": qualidade}

    template = (Path(__file__).parent / "template.html").read_text(encoding="utf-8")
    html = template.replace("/*__DADOS__*/null", json.dumps(dados, ensure_ascii=False))
    Path(saida).write_text(html, encoding="utf-8")
    print(f"OK: {saida}  ({len(lojas)} lojas, {len(linhas)} linhas de export em {len(oficina)} meses, {len(estoque)} linhas de estoque)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "index.html")
