"""
Simulador IGEST, versão 2: reimplementação e revisão da metodologia.

Reproduz exatamente o cálculo do app Shiny original (versão 1) e o compara com
a versão revisada, que segue o Manual de Uso do IGEST (CGJT) literalmente.

Uso:  python igest_v2.py            -> imprime o diagnóstico da revisão
      python igest_v2.py --export   -> grava ../site/igest/dados.js
Requer: openpyxl
"""
from __future__ import annotations

import json
import math
import sys
from pathlib import Path

import openpyxl

AQUI = Path(__file__).resolve().parent
ARQUIVO = AQUI / "arquivo.xlsx"

# Linhas (1-based, sem cabeçalho) de cada TRT na planilha, como no app original
LIMITES_TRT = [1, 147, 364, 522, 654, 742, 809, 846, 902, 998, 1033, 1065, 1125,
               1152, 1184, 1337, 1360, 1384, 1432, 1454, 1469, 1492, 1506, 1544, 1570]
SIGLAS_TRT = ["RJ", "SP", "MG", "RS", "BA", "PE", "CE", "PA/AP", "PR", "DF/TO", "AM/RR", "SC",
              "PB", "RO/AC", "Campinas", "MA", "ES", "GO", "AL", "SE", "RN", "PI", "MT", "MS"]

PESOS = {  # manual, item 7
    "Acervo": {"I01": .3, "I02": .4, "I03": .3},
    "Celeridade": {"I04": .6, "I05": .1, "I06": .3},
    "Produtividade": {"I07": .3, "I08": .4, "I13": .3},
    "Congestionamento": {"I09": .5, "I10": .5},
    "Força de trabalho": {"I11": .5, "I12": .5},
}
POSITIVOS = {"I07", "I08", "I11", "I13"}  # manual, item 9
INDICADORES = [f"I{i:02d}" for i in range(1, 14)]


# ----------------------------------------------------------------------------- dados
def carrega():
    ws = openpyxl.load_workbook(ARQUIVO, read_only=True, data_only=True).worksheets[0]
    linhas = list(ws.iter_rows(values_only=True))
    cab, dados = linhas[0], linhas[1:]
    varas = []
    for i, r in enumerate(dados, start=1):
        trt = next(t for t in range(24) if LIMITES_TRT[t] <= i < LIMITES_TRT[t + 1]) + 1
        varas.append({"id": i - 1, "nome": " ".join(r[0].split()), "trt": trt,
                      "v": [float(x) for x in r[1:18]]})
    return cab, varas


# ----------------------------------------------------------------------- indicadores
def _div(a, b):
    return a / b if b != 0 else None  # indefinido


def indicadores(v):
    """13 indicadores brutos a partir de V01..V17 (manual, glossário). None = indefinido."""
    V = [None] + list(v)  # V[1]..V[17]
    pct = lambda a, b: None if _div(a, b) is None else 100 * a / b
    return {
        "I01": V[1],
        "I02": V[2] + V[3] + V[4],
        "I03": pct(V[11], V[12]),
        "I04": V[5], "I05": V[6], "I06": V[7],
        "I07": pct(V[8], V[9]),
        "I08": pct(V[9], V[10]),
        "I09": pct(V[2], V[13] + V[2]),
        "I10": pct(V[4], V[14] + V[4]),
        "I11": _div(V[13] + V[14], V[15]),
        "I12": _div(V[2] + V[4], V[15]),
        "I13": pct(V[17], V[16]),
    }


def quantil(xs, p):
    """Quantil tipo 7 (padrão do R e do numpy)."""
    s = sorted(xs)
    h = (len(s) - 1) * p
    lo = math.floor(h)
    hi = min(lo + 1, len(s) - 1)
    return s[lo] + (h - lo) * (s[hi] - s[lo])


def limites(valores, truncar_li):
    q1, q3 = quantil(valores, .25), quantil(valores, .75)
    li, ls = q1 - 1.5 * (q3 - q1), q3 + 1.5 * (q3 - q1)
    if truncar_li and li < 0:
        li = 0.0
    return li, ls


def pontua(x, li, ls, positivo):
    if ls == li:
        z = 0.0
    elif x < li:
        z = 0.0
    elif x > ls:
        z = 1.0
    else:
        z = (x - li) / (ls - li)
    return 1 - z if positivo else z


# --------------------------------------------------------------------------- motores
def calcula(grupo, versao):
    """grupo: lista de vetores V (17). Retorna lista de dicts com escores e IGEST.

    versao 1 = app Shiny original: indefinido vira 0, Inf continua Inf, LI truncado em 0.
    versao 2 = manual literal: LI sem truncar, indefinido fica fora dos limites e o
               mesoindicador é reponderado pelos indicadores disponíveis.
    """
    brutos = [indicadores(v) for v in grupo]
    if versao == 1:
        for b, v in zip(brutos, grupo):
            for k, x in b.items():
                if x is None:
                    # 0/0 -> NaN -> 0 no R; x/0 -> Inf (fica Inf)
                    b[k] = 0.0 if _nan_no_r(k, b, v) else math.inf
    lim = {}
    for k in INDICADORES:
        xs = [b[k] for b in brutos if b[k] is not None and math.isfinite(b[k])] if versao == 2 \
            else [b[k] for b in brutos]
        lim[k] = limites(xs, truncar_li=(versao == 1))
    saida = []
    for b in brutos:
        esc = {k: (None if b[k] is None else pontua(b[k], *lim[k], k in POSITIVOS)) for k in INDICADORES}
        meso = {}
        for m, pesos in PESOS.items():
            disp = {k: w for k, w in pesos.items() if esc[k] is not None}
            tot = sum(disp.values())
            meso[m] = sum(esc[k] * w for k, w in disp.items()) / tot if tot else None
        ig = sum(.2 * x for x in meso.values() if x is not None)
        saida.append({"bruto": b, "esc": esc, "meso": meso, "igest": ig})
    return saida, lim


def _nan_no_r(k, b, v):
    """No R, 0/0 = NaN (vira 0 pelo is.na) e x/0 = Inf (não é NA)."""
    num = {"I03": v[10], "I07": v[7], "I08": v[8], "I09": v[1], "I10": v[3],
           "I11": v[12] + v[13], "I12": v[1] + v[3], "I13": v[16]}[k]
    return num == 0


def ranking(res):
    ordem = sorted(range(len(res)), key=lambda i: (res[i]["igest"], i))
    pos = [0] * len(res)
    for p, i in enumerate(ordem, start=1):
        pos[i] = p
    return pos


# ----------------------------------------------------------------- diagnóstico v1×v2
def faixa(recebidos):
    for i, teto in enumerate([500, 750, 1000, 1500, 2000, 2500]):
        if recebidos <= teto:
            return i
    return 6


def diagnostico():
    cab, varas = carrega()
    V = [x["v"] for x in varas]
    n = len(varas)
    print(f"varas: {n}")

    # 1) dados
    inteiros = sum(abs(x - round(x)) < 1e-9 for v in V for x in v)
    print(f"valores inteiros nas 17 variáveis: {inteiros} de {n*17}")
    print(f"V15 (servidores) min/máx: {min(v[14] for v in V):.3f} / {max(v[14] for v in V):.1f}")
    nomes = {}
    for x in varas:
        nomes.setdefault(x["nome"], []).append(x["trt"])
    print("nomes repetidos:", {k: t for k, t in nomes.items() if len(t) > 1})

    # 2) indefinidos
    for k in INDICADORES:
        ind = sum(indicadores(v)[k] is None for v in V)
        if ind:
            print(f"{k}: {ind} varas com denominador zero")
    v12z = [v for v in V if v[11] == 0]
    print(f"  dessas, V11=0 (0/0 -> nota máxima na v1): {sum(v[10]==0 for v in v12z)}; "
          f"V11>0 (x/0 -> Inf -> nota mínima): {sum(v[10]>0 for v in v12z)}")

    # 3) limites v1 vs v2
    r1, l1 = calcula(V, 1)
    r2, l2 = calcula(V, 2)
    print("\nlimites (nacional)   LI v1 | LI v2 (manual) | LS")
    for k in INDICADORES:
        print(f"  {k}: {l1[k][0]:10.3f} | {l2[k][0]:10.3f} | {l2[k][1]:10.3f}")
    trunc = [k for k in INDICADORES if l2[k][0] < 0]
    print("indicadores com LI negativo (truncado na v1):", trunc)

    p1, p2 = ranking(r1), ranking(r2)
    d = [abs(a - b) for a, b in zip(p1, p2)]
    dig = [abs(a["igest"] - b["igest"]) for a, b in zip(r1, r2)]
    print(f"\nIGEST v1×v2: dif. média {sum(dig)/n:.4f}, máx {max(dig):.4f}")
    print(f"posição v1×v2: mediana {sorted(d)[n//2]}, p90 {sorted(d)[int(n*.9)]}, máx {max(d)}; "
          f"mudaram >50 posições: {sum(x>50 for x in d)}")
    q = lambda p: 0 if p <= n * .25 else (2 if p > n * .75 else 1)
    troca = sum(q(a) != q(b) for a, b in zip(p1, p2))
    print(f"varas que mudam de faixa (25% melhores / 50% / 25% piores): {troca}")

    # 4) TRT10
    print("\nTRT10 (DF/TO) — posição nacional v1 -> v2")
    for i, x in enumerate(varas):
        if x["trt"] == 10:
            print(f"  {x['nome']:28s} {p1[i]:5d} -> {p2[i]:5d}   IGEST {r1[i]['igest']:.4f} -> {r2[i]['igest']:.4f}")

    # 5) bug do grepl: Palmas
    pal = [i for i, x in enumerate(varas) if x["nome"] == "Palmas - 01a Vara"]
    print("\nPalmas - 01a Vara aparece em:", [varas[i]["trt"] for i in pal],
          "-> no comparativo nacional a v1 remove as duas (grepl)")

    # 6) faixas
    cont = [0] * 7
    for v in V:
        cont[faixa(v[9])] += 1
    print("\nvaras por faixa de recebidos (V10):", cont)

    # 7) sensibilidade nacional: melhorar 10% cada variável, efeito médio no IGEST da TRT10
    return varas, r1, r2, l1, l2, p1, p2


def exporta():
    _, varas = carrega()
    destino = AQUI.parent / "site" / "igest" / "dados.js"
    destino.parent.mkdir(parents=True, exist_ok=True)
    out = {"trt": SIGLAS_TRT,
           "varas": [[x["nome"], x["trt"], [round(y, 4) for y in x["v"]]] for x in varas]}
    js = "window.IGEST_DATA=" + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";\n"
    destino.write_text(js, encoding="utf-8")
    print("gravado", destino, destino.stat().st_size, "bytes")


if __name__ == "__main__":
    if "--export" in sys.argv:
        exporta()
    else:
        diagnostico()
