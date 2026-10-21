"""Az egyedi Harsányi-osztalékok (ψ({R}), ψ({S}), ψ({C})) pontos bontása régió-, szektor- és részvényszintű hozzájárulásokra.

Miért kell? A ψ({R}) = −58,07 százalékpont azt mondja meg, hogy a régiósúlyok eltérése összesen mennyit hozott, de azt nem, hogy
melyik régió súlyeltérése járult hozzá ehhez és mennyivel. Ez a szkript megadja, melyik régió, szektor, illetve részvény mennyit
tett hozzá az egyes egyedi osztalékokhoz (kumulált százalékpontban, a Davies–Laker-féle értékfüggvénnyel összhangban).

Módszer (pontos, nem közelítés): két hozamsorozat kumulált különbsége
    Π(1 + a_t) − Π(1 + b_t) = Σ_t [Π_{s<t}(1 + a_s)] · (a_t − b_t) · [Π_{s>t}(1 + b_s)]
és az (a_t − b_t) havi különbség additív a régiók (szektorok, részvények) havi hozzájárulásaiban, ezért minden szereplő
hozzájárulása a teljes kumulált különbségből pontosan megkapható; a hozzájárulások összege egyenlő az egyedi osztalékkal
(a szkript ezt ellenőrzi).

A két "véglegeshozamszámítás" notébook kódját használja (nem módosítja őket, a kimeneti xlsx-eket nem írja felül).
Futtatás:  python hozzajarulas_bontas.py      (a "VS code" mappából vagy a projekt gyökeréből)
Kimenet:   hozzajarulas_bontas_eredmenyek.xlsx (lapok: <stratégia>_régió, <stratégia>_szektor, <stratégia>_részvény)
Időszak:   a teljes időszak (2016–2025, benchmark-referencia: 2015.12.31.)
"""
import json
import os
import sys
import types
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
os.chdir(HERE)

NOTEBOOKS = {"EW": "véglegeshozamszámítás.ipynb", "MOM": "véglegeshozamszámítás_momentum.ipynb"}
NAMES = {"EW": "egyenlő_súly", "MOM": "momentum"}
PSI_SHEET = {"EW": "outputs_harsanyi_100stocks_hierarchical_davies_laker", "MOM": "outputs_harsanyi_100stocks_momentum"}


def load_model(nb_name: str, mod_name: str) -> types.ModuleType:
    nb = json.load(open(HERE / nb_name, encoding="utf8"))
    code = "\n".join("".join(c["source"]) for c in nb["cells"] if c["cell_type"] == "code")
    mod = types.ModuleType(mod_name)
    mod.__dict__["__name__"] = mod_name
    sys.modules[mod_name] = mod
    exec(compile(code, nb_name, "exec"), mod.__dict__)
    return mod


def exact_contributions(c_df: pd.DataFrame, b: pd.Series):
    """c_df: hónap × szereplő (a havi hozamkülönbség additív részei); b: a benchmark havi hozama."""
    a = b + c_df.sum(axis=1)
    p = (1 + a).cumprod().shift(1).fillna(1.0)
    q = (1 + b)[::-1].cumprod()[::-1].shift(-1).fillna(1.0)
    contrib = c_df.mul(p * q, axis=0).sum()
    total = float((1 + a).prod() - (1 + b).prod())
    return contrib, total


writer = pd.ExcelWriter(HERE / "hozzajarulas_bontas_eredmenyek.xlsx", engine="openpyxl")
for key, nb in NOTEBOOKS.items():
    mod = load_model(nb, "hb_" + key)
    data = mod.load_portfolio_data()
    d = mod.slice_data_by_period(data, "2016-01-01", "2025-12-31", mod.PERIOD_BENCHMARK_REFERENCE_DATES["full_2016_2025"])
    R, meta = d.returns, d.metadata
    act = mod.build_attribution_actions(R, d.active_weights, meta)
    ben = mod.build_attribution_actions(R, d.benchmark_weights, meta)

    def time_varying(a):
        rw, sw, cr = a.region_weights, a.sector_weights, a.cell_returns
        if isinstance(rw, pd.Series):
            rw = pd.DataFrame([rw.values] * len(R), index=R.index, columns=rw.index)
        if isinstance(sw, pd.Series):
            sw = pd.DataFrame([sw.values] * len(R), index=R.index, columns=sw.index)
        return rw, sw, cr

    rwA, swA, crA = time_varying(act)
    rwB, swB, crB = time_varying(ben)
    b = sum(rwB[r] * swB[(r, s)] * crB[(r, s)] for (r, s) in crB.columns)

    comp_r = {r: sum(swB[c] * crB[c] for c in crB.columns if c[0] == r) for r in rwB.columns}
    reg = pd.DataFrame({r: (rwA[r] - rwB[r]) * comp_r[r] for r in rwB.columns})
    reg_c, v_reg = exact_contributions(reg, b)
    sd = {}
    for (r, s) in crB.columns:
        sd[s] = sd.get(s, 0) + rwB[r] * (swA[(r, s)] - swB[(r, s)]) * crB[(r, s)]
    sec_c, v_sec = exact_contributions(pd.DataFrame(sd), b)
    wA = d.active_weights
    if isinstance(wA, pd.Series):
        wA = pd.DataFrame([wA.values] * len(R), index=R.index, columns=wA.index)
    wB = pd.DataFrame([d.benchmark_weights.values] * len(R), index=R.index, columns=d.benchmark_weights.index)
    cs = {}
    for (r, s), g in meta.loc[R.columns].groupby(["region", "sector"]):
        tk = list(g.index)
        bt = wB[tk].sum(axis=1)
        an = wA[tk].div(wA[tk].sum(axis=1), axis=0)
        for t in tk:
            cs[t] = (bt * an[t] - wB[t]) * R[t]
    sel_c, v_sel = exact_contributions(pd.DataFrame(cs), b)

    for nm, got, want in (("régió", reg_c.sum(), v_reg), ("szektor", sec_c.sum(), v_sec), ("szelekció", sel_c.sum(), v_sel)):
        assert abs(got - want) < 1e-9, (key, nm, got, want)
    print("%s: a hozzájárulások összege = az egyedi osztalék (régió %.2f, szektor %.2f, szelekció %.2f százalékpont)" % (
        NAMES[key], 100 * v_reg, 100 * v_sec, 100 * v_sel))

    name = meta["company"] if "company" in meta.columns else pd.Series(meta.index, index=meta.index)
    (100 * reg_c).rename("hozzájárulás (százalékpont)").sort_values().to_frame().to_excel(writer, sheet_name=NAMES[key] + "_régió")
    (100 * sec_c).rename("hozzájárulás (százalékpont)").sort_values().to_frame().to_excel(writer, sheet_name=NAMES[key] + "_szektor")
    stock = pd.DataFrame({"vállalat": name.loc[sel_c.index], "hozzájárulás (százalékpont)": 100 * sel_c}).sort_values("hozzájárulás (százalékpont)")
    stock.to_excel(writer, sheet_name=NAMES[key] + "_részvény")
writer.close()
print("kész: hozzajarulas_bontas_eredmenyek.xlsx")
