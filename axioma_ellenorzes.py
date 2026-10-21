"""A Harsányi-attribúció három axiómájának számszerű ellenőrzése a dolgozat mintáján.

A szkript a két "véglegeshozamszámítás" notébook kódját használja (nem módosítja őket, és a kimeneti
xlsx-eket sem írja felül), és mindkét aktív stratégiára (egyenlő súlyozás, momentum), a teljes időszakra és a három
részidőszakra lefuttatja az ellenőrzéseket (Acerbi, Csóka és Herings, 2026):

  1. Hatékonyság: a hét Harsányi-osztalék összege egyenlő a v(N) értékkel (a Shapley-értékek összege is).
  2. Külső döntéshozótól való függetlenség: ha egy döntéshozó cselekvését megváltoztatjuk (az aktív és a benchmark
     cselekvés 50-50%-os keverékére), a hozzá nem tartozó koalíciók osztaléka nem változik.
  3. Nulla hozzájárulás: ha egy döntéshozó a benchmark cselekvést választja, az őt tartalmazó koalíciók osztaléka nulla.

Ezen kívül összeveti az újraszámolt osztalékokat a notébookok által mentett xlsx-ek értékeivel.

Futtatás (a projekt gyökeréből vagy a "VS code" mappából):
    python axioma_ellenorzes.py
Kimenet: a konzolra írt összegzés és az axioma_ellenorzes_eredmenyek.xlsx ugyanebben a mappában.
"""
import glob
import itertools
import json
import os
import sys
import types
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
os.chdir(HERE)

NOTEBOOKS = {
    "Egyenlő súlyozás": ("véglegeshozamszámítás.ipynb", "outputs_harsanyi_100stocks_hierarchical_davies_laker"),
    "Momentum": ("véglegeshozamszámítás_momentum.ipynb", "outputs_harsanyi_100stocks_momentum"),
}
PERIODS = {
    "full_2016_2025": ("2016-01-01", "2025-12-31"),
    "pre_covid_2016_2019": ("2016-01-01", "2019-12-31"),
    "covid_2020_2021": ("2020-01-01", "2021-12-31"),
    "inflation_ai_2022_2025": ("2022-01-01", "2025-12-31"),
}
PLAYERS = ["REGION", "SECTOR", "SELECTION"]
XLSX_PREFIX = {"full_2016_2025": "full_2016_2025", "pre_covid_2016_2019": "pre_covid_2016_2019",
               "covid_2020_2021": "covid_2020_2021", "inflation_ai_2022_2025": "infl_ai_2022_2025"}


def load_model(nb_name: str, mod_name: str) -> types.ModuleType:
    """A notébook kódcelláit modulként tölti be (a main() nem fut le)."""
    nb = json.load(open(HERE / nb_name, encoding="utf8"))
    code = "\n".join("".join(c["source"]) for c in nb["cells"] if c["cell_type"] == "code")
    mod = types.ModuleType(mod_name)
    mod.__dict__["__name__"] = mod_name
    sys.modules[mod_name] = mod
    exec(compile(code, nb_name, "exec"), mod.__dict__)
    return mod


def blend(a, b, w=0.5):
    """Két (azonos szerkezetű) pandas objektum súlyozott keveréke."""
    return (a * w).add(b * (1.0 - w), fill_value=0.0)


def replace_action(mod, actions, player, new_actions):
    kw = dict(region_weights=actions.region_weights, sector_weights=actions.sector_weights, cell_returns=actions.cell_returns)
    key = {"REGION": "region_weights", "SECTOR": "sector_weights", "SELECTION": "cell_returns"}[player]
    kw[key] = getattr(new_actions, key)
    return mod.AttributionActions(**kw)


def dividends(mod, active, bench, all_cells):
    values = mod.compute_coalition_values(active, bench, PLAYERS, all_cells)
    return values, mod.harsanyi_attributions(values, PLAYERS)


def label(k):
    return "{" + ",".join({"REGION": "R", "SECTOR": "S", "SELECTION": "C"}[p] for p in PLAYERS if p in k) + "}"


rows = []
for strategy, (nb_name, out_dir) in NOTEBOOKS.items():
    mod = load_model(nb_name, "harsanyi_" + ("ew" if strategy.startswith("Egy") else "mom"))
    data = mod.load_portfolio_data()
    xlsx = glob.glob(str(HERE / out_dir / "*.xlsx"))[0]
    for period, (start, end) in PERIODS.items():
        d = mod.slice_data_by_period(data, start, end, mod.PERIOD_BENCHMARK_REFERENCE_DATES[period])
        active = mod.build_attribution_actions(d.returns, d.active_weights, d.metadata)
        bench = mod.build_attribution_actions(d.returns, d.benchmark_weights, d.metadata)
        cells = bench.cell_returns.columns.union(active.cell_returns.columns)
        values, psi = dividends(mod, active, bench, cells)
        vN = values[frozenset(PLAYERS)]
        shap = mod.shapley_from_harsanyi(psi, PLAYERS)["shapley_value"]

        saved = pd.read_excel(xlsx, sheet_name=XLSX_PREFIX[period] + "_attr")
        name = {"REGION": "R", "SECTOR": "S", "SELECTION": "C"}
        dev = 0.0
        for r in saved.itertuples():
            k = frozenset(x.strip() for x in r.coalition.split("+"))
            dev = max(dev, abs(psi[k] - r.harsanyi_attribution))
        rows.append((strategy, period, "0. egyezés a mentett xlsx-szel", "mind a hét osztalék", dev))

        rows.append((strategy, period, "1. hatékonyság", "Σψ − v(N)", abs(sum(psi.values()) - vN)))
        rows.append((strategy, period, "1. hatékonyság", "Σφ − v(N)", abs(float(shap.sum()) - vN)))

        for x in PLAYERS:
            active_null = replace_action(mod, active, x, bench)
            _, psi_null = dividends(mod, active_null, bench, cells)
            dev3 = max(abs(v) for k, v in psi_null.items() if x in k)
            rows.append((strategy, period, "3. nulla hozzájárulás", "%s a benchmark cselekvést választja" % x, dev3))

            changed = mod.AttributionActions(
                region_weights=blend(active.region_weights, bench.region_weights) if x == "REGION" else active.region_weights,
                sector_weights=blend(active.sector_weights, bench.sector_weights) if x == "SECTOR" else active.sector_weights,
                cell_returns=blend(active.cell_returns, bench.cell_returns) if x == "SELECTION" else active.cell_returns,
            )
            values_c, psi_c = dividends(mod, changed, bench, cells)
            dev2 = max(abs(psi_c[k] - psi[k]) for k in psi if x not in k)
            moved = max(abs(psi_c[k] - psi[k]) for k in psi if x in k)
            rows.append((strategy, period, "2. külső döntéshozótól való függetlenség", "%s cselekvése megváltozik" % x, dev2))
            rows.append((strategy, period, "2. kontroll: az érintett koalíciók változnak", "%s cselekvése megváltozik" % x, moved))

res = pd.DataFrame(rows, columns=["stratégia", "időszak", "axióma", "teszt", "legnagyobb eltérés (hozam-egység)"])
res.to_excel(HERE / "axioma_ellenorzes_eredmenyek.xlsx", index=False)

pd.set_option("display.width", 220, "display.max_colwidth", 60)
summary = (res.groupby(["stratégia", "axióma"])["legnagyobb eltérés (hozam-egység)"].max().reset_index())
print(summary.to_string(index=False, float_format=lambda v: "%.3e" % v))
print("\nA kontroll-sorokban (az érintett koalíciók változása) a legkisebb változás a teszt érzékenységét mutatja:")
ctl = res[res["axióma"].str.startswith("2. kontroll")]
print("  legkisebb változás: %.3e (százalékpontban: %.4f)" % (ctl["legnagyobb eltérés (hozam-egység)"].min(), 100 * ctl["legnagyobb eltérés (hozam-egység)"].min()))
