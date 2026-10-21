"""
Az adatelemzési fejezet ábrái (3-7. ábra) a Harsányi-attribúció kimeneti Excel-fájljaiból.

Bemenet (ebben a mappában, a három notebook kimenete):
    outputs_harsanyi_100stocks_hierarchical_davies_laker\\*.xlsx   egyenlő súlyozású aktív portfólió
    outputs_harsanyi_100stocks_momentum\\*.xlsx                      momentum-stratégiájú aktív portfólió
    adatok_python.xlsx                                              csak a benchmark régiósúlyaihoz
                                                                    (a 2019.12.31. és 2021.12.31. referencia-dátumon)

Kimenet (alapértelmezés szerint a Szakdolgozat\\ábrák mappa; PNG 300 dpi és SVG):
    abra_momentum_regiosulyok     3. ábra  A momentum-portfólió régiósúlya időben
    abra_kumulalt_tobblethozam    4. ábra  A kumulált többlethozam alakulása
    abra_osztalek_vizeses_egyenlo   5. ábra  A hét Harsányi-osztalék vízesés-ábrája, egyenlő súlyozás
    abra_osztalek_vizeses_momentum  6. ábra  A hét Harsányi-osztalék vízesés-ábrája, momentum
    abra_osztalekok_strategiak      7. ábra  A hét Harsányi-osztalék időszakonként, a két stratégiára
    abra_shapley_strategiak         (a Shapley-értékek időszakonként; a fejezetben már nem szerepel)

Futtatás (a notebookok újrafuttatása után ezt is újra kell futtatni):
    python abrak_keszitese.py                  # kimenet: ..\\ábrák
    python abrak_keszitese.py --out C:\\masik\\mappa

Szükséges csomagok: pandas, numpy, matplotlib (3.6 vagy újabb), openpyxl. A betűtípus Times New Roman (ha nincs, DejaVu Serif).
Az ábrák csak a fenti Excel-fájlokból számolnak; a számok nincsenek beleírva a kódba, a futás elején pedig
a szkript ellenőrzi, hogy a hatékonysági azonosság (az osztalékok összege = a teljes többlethozam) teljesül.
"""
from __future__ import annotations

import argparse
import glob
import sys
from pathlib import Path

import matplotlib as mpl

mpl.use("Agg")
import matplotlib.dates as mdates
import matplotlib.patheffects as pe
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from matplotlib.ticker import FuncFormatter, MultipleLocator

BASE = Path(__file__).resolve().parent
DEFAULT_OUT = BASE.parent / "ábrák"

PATTERN_EW = "outputs_*davies_laker/*.xlsx"
PATTERN_MOM = "outputs_*momentum/*.xlsx"

PERIODS = [
    ("full_2016_2025", "2016–2025", "teljes időszak"),
    ("pre_covid_2016_2019", "2016–2019", "Covid előtt"),
    ("covid_2020_2021", "2020–2021", "Covid"),
    ("infl_ai_2022_2025", "2022–2025", "infláció–MI"),
]
FULL, SUBPERIODS = PERIODS[0][0], [p[0] for p in PERIODS[1:]]
TITLE = {k: f"{short} ({desc})" for k, short, desc in PERIODS}
SHORT = {k: short for k, short, _ in PERIODS}

ORDER = ["REGION", "SECTOR", "SELECTION", "REGION + SECTOR", "REGION + SELECTION", "SECTOR + SELECTION", "REGION + SECTOR + SELECTION"]
XLABELS = ["{R}", "{S}", "{C}", "{R,S}", "{R,C}", "{S,C}", "{R,S,C}", "Összesen"]
DECISION_MAKERS = ["REGION", "SECTOR", "SELECTION"]

SURFACE = "#ffffff"
INK, INK2, MUTED = "#0b0b0b", "#52514e", "#898781"
GRID, BASELINE = "#e1e0d9", "#c3c2b7"
C_EW, C_MOM = "#2a78d6", "#eb6834"
C_NA, C_EU, C_AS = "#1baf7a", "#4a3aa7", "#e87ba4"
BAND = "#f3f2ee"

FONT_CANDIDATES = ["Times New Roman", "Liberation Serif", "DejaVu Serif"]


def find_xlsx(pattern: str) -> Path:
    files = sorted(glob.glob(str(BASE / pattern)))
    if not files:
        sys.exit(f"HIBA: nem találok kimeneti fájlt erre a mintára: {BASE / pattern}\n"
                 "Futtasd le előbb a megfelelő notebookot.")
    if len(files) > 1:
        print(f"Figyelem: több fájl is van ({pattern}); a következőt használom: {files[-1]}")
    return Path(files[-1])


def load_strategy(path: Path) -> dict:
    """Egy stratégia kimeneti munkafüzete -> időszakonkénti szótár."""
    x = pd.ExcelFile(path)
    out = {"path": path, "periods": {}}
    for key, _, _ in PERIODS:
        attr = x.parse(f"{key}_attr").set_index("coalition")["harsanyi_attribution_pct_point"]
        shap = x.parse(f"{key}_shapley").set_index("decision_maker")["shapley_value_pct_point"]
        eff = x.parse(f"{key}_eff").iloc[0]
        ts = x.parse(f"{key}_ts", index_col=0)
        v_n = float(eff["total_excess_vN_pct_point"])
        assert abs(attr.sum() - v_n) < 1e-6, f"{path.name}: {key}: az osztalékok összege nem egyezik v(N)-nel"
        assert abs(shap.sum() - v_n) < 1e-6, f"{path.name}: {key}: a Shapley-értékek összege nem egyezik v(N)-nel"
        out["periods"][key] = {
            "attr": attr, "shapley": shap, "vN": v_n, "ts": ts,
            "start": pd.Timestamp(eff["start_date"]), "end": pd.Timestamp(eff["end_date"]),
            "ref": pd.Timestamp(eff["benchmark_reference_date"]),
        }
    out["stock_weights"] = x.parse("stock_weights")
    out["region_weights"] = x.parse("region_weights")
    if "active_region_weights_t" in x.sheet_names:
        out["region_weights_t"] = x.parse("active_region_weights_t", index_col=0)
    out["input_xlsx"] = BASE / str(x.parse("data_quality").set_index("metric").loc["excel_input_path", "value"])
    return out


def benchmark_na_weights(ew: dict) -> dict:
    """Az Észak-Amerika benchmark-súlya (%) minden elemzési időszak referencia-dátumán (kapitalizáció alapján)."""
    market_cap = pd.read_excel(ew["input_xlsx"], sheet_name="Market_Cap", index_col=0)
    region = ew["stock_weights"].set_index("ticker")["region"]
    out = {}
    for key in SUBPERIODS:
        ref = ew["periods"][key]["ref"]
        w = market_cap.loc[ref]
        w = w / w.sum()
        out[key] = float(w.groupby(region.reindex(w.index)).sum()["North America"] * 100)
    rw = ew["region_weights"].set_index("region")["benchmark_region_weight"]["North America"] * 100
    assert abs(out[SUBPERIODS[0]] - rw) < 1e-6, "a benchmark régiósúly nem egyezik a region_weights lappal"
    return out


def hu(x, d=2, plus=False) -> str:
    """Magyar tizedesvessző és valódi mínuszjel."""
    r = round(float(x), d)
    s = f"{abs(r):.{d}f}".replace(".", ",")
    return "\u2212" + s if r < 0 else (("+" + s) if plus and r > 0 else s)


FMT0 = FuncFormatter(lambda v, p: hu(v, 0))


def setup_style() -> None:
    from matplotlib import font_manager as fm
    installed = {f.name for f in fm.fontManager.ttflist}
    family = next((f for f in FONT_CANDIDATES if f in installed), "DejaVu Serif")
    if family != FONT_CANDIDATES[0]:
        print(f"Figyelem: a {FONT_CANDIDATES[0]} betűtípus nem található, helyette: {family}")
    mpl.rcParams.update({
        "font.family": "serif", "font.serif": [family], "font.size": 9,
        "axes.unicode_minus": True, "svg.fonttype": "path",
        "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
        "axes.edgecolor": BASELINE, "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2, "text.color": INK,
        "axes.linewidth": 0.8, "xtick.major.width": 0.8, "xtick.major.size": 3, "ytick.major.width": 0.0, "ytick.major.size": 0,
        "axes.spines.top": False, "axes.spines.right": False, "axes.spines.left": False, "legend.frameon": False,
    })


def style_axes(ax, grid: bool = True) -> None:
    ax.spines["bottom"].set_color(BASELINE)
    if grid:
        ax.yaxis.grid(True, color=GRID, lw=0.6, zorder=0)
        ax.set_axisbelow(True)
    ax.tick_params(axis="x", labelsize=8.5)
    ax.tick_params(axis="y", labelsize=8.5)


def save(fig, out_dir: Path, name: str) -> None:
    for ext in ("png", "svg"):
        fig.savefig(out_dir / f"{name}.{ext}", dpi=300 if ext == "png" else None, bbox_inches="tight", pad_inches=0.06)
    plt.close(fig)
    print(f"  mentve: {out_dir / name}.png / .svg")


STRATEGY_LINES = [Line2D([0], [0], color=C_EW, lw=2.2, label="Egyenlő súlyozás"), Line2D([0], [0], color=C_MOM, lw=2.2, label="Momentum")]
STRATEGY_PATCHES = [Patch(facecolor=C_EW, label="Egyenlő súlyozás"), Patch(facecolor=C_MOM, label="Momentum")]


def fig_region_weights(ew: dict, mom: dict, bench_na: dict, out_dir: Path) -> None:
    """3. ábra: a momentum-portfólió aktív régiósúlya időben (halmozott terület)."""
    w = mom["region_weights_t"] * 100
    first, last_day = ew["periods"][FULL]["start"], ew["periods"][FULL]["end"] + pd.Timedelta(days=1)
    w = pd.concat([w.iloc[[0]].set_axis([first]), w, w.iloc[[-1]].set_axis([last_day])])
    w = w[~w.index.duplicated(keep="last")]

    fig, ax = plt.subplots(figsize=(6.3, 3.7))
    fig.subplots_adjust(left=0.10, right=0.82, top=0.95, bottom=0.26)
    ax.stackplot(w.index, w["North America"], w["Europe"], w["Asia"], colors=[C_NA, C_EU, C_AS], step="post",
                 edgecolors=SURFACE, linewidth=1.0, zorder=2)
    ax.set_xlim(first, last_day)
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_locator(MultipleLocator(20))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: hu(v, 0) + "%"))
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    style_axes(ax, grid=False)
    ax.spines["left"].set_visible(True)
    ax.spines["left"].set_color(BASELINE)

    halo = [pe.withStroke(linewidth=3.2, foreground=SURFACE)]
    for key in SUBPERIODS:
        p = ew["periods"][key]
        x0, x1 = p["start"], p["end"] + pd.Timedelta(days=1)
        v = bench_na[key]
        ax.plot([x0, x1], [v, v], color=INK, lw=1.6, solid_capstyle="butt", path_effects=halo, zorder=5)
        ax.text(x0 + (x1 - x0) / 2, v + 2.2, hu(v, 2) + "%", ha="center", va="bottom", fontsize=8.5, color=INK, path_effects=halo, zorder=6)
    ew_w = float(ew["region_weights"]["active_region_weight"].iloc[0] * 100)
    ax.axhline(ew_w, color=INK, lw=1.0, path_effects=halo, zorder=5)

    last = w.iloc[-1]
    ys = [last["North America"] / 2, last["North America"] + last["Europe"] / 2, last["North America"] + last["Europe"] + last["Asia"] / 2]
    for y, t in zip(ys, ("Észak-Amerika", "Európa", "Ázsia")):
        ax.annotate(t, xy=(last_day, y), xytext=(6, 0), textcoords="offset points", va="center", fontsize=8.5, color=INK, annotation_clip=False)
    ax.set_ylabel("Aktív régiósúly", fontsize=9)
    areas = [Patch(facecolor=C_NA, label="Észak-Amerika"), Patch(facecolor=C_EU, label="Európa"), Patch(facecolor=C_AS, label="Ázsia")]
    lines = [Line2D([0], [0], color=INK, lw=1.6, label="Benchmark: az Észak-Amerika súlya az adott időszakban"),
             Line2D([0], [0], color=INK, lw=1.0, label=f"Egyenlő súlyozás: {hu(ew_w, 1)}% régiónként")]
    fig.legend(handles=areas, loc="lower left", ncol=3, fontsize=8.5, bbox_to_anchor=(0.10, 0.085), handlelength=1.6, columnspacing=1.6,
               title="Momentum-portfólió régiósúlya", title_fontsize=8.5, alignment="left")
    fig.legend(handles=lines, loc="lower left", ncol=1, fontsize=8.5, bbox_to_anchor=(0.10, -0.02), handlelength=1.8)
    save(fig, out_dir, "abra_momentum_regiosulyok")


def fig_cumulative_excess(ew: dict, mom: dict, out_dir: Path) -> None:
    """4. ábra: a kumulált többlethozam időben, a két stratégia egy ábrán."""
    ce = ew["periods"][FULL]["ts"]["cumulative_excess_return"] * 100
    cm = mom["periods"][FULL]["ts"]["cumulative_excess_return"] * 100
    x0 = ew["periods"][FULL]["start"]
    x1 = ew["periods"][FULL]["end"] + pd.Timedelta(days=1)

    fig, ax = plt.subplots(figsize=(6.3, 3.5))
    fig.subplots_adjust(left=0.10, right=0.80, top=0.93, bottom=0.17)
    cov = ew["periods"]["covid_2020_2021"]
    c0, c1 = cov["start"], cov["end"] + pd.Timedelta(days=1)
    ax.axvspan(c0, c1, color=BAND, lw=0, zorder=0)
    for xv in (c0, c1):
        ax.axvline(xv, color=BASELINE, lw=0.8, zorder=1)
    ax.axhline(0, color=MUTED, lw=0.9, zorder=2)
    ax.plot(ce.index, ce.values, color=C_EW, lw=1.7, zorder=3, solid_capstyle="round")
    ax.plot(cm.index, cm.values, color=C_MOM, lw=1.7, zorder=3, solid_capstyle="round")
    for s, c in ((ce, C_EW), (cm, C_MOM)):
        ax.plot(s.index[-1], s.iloc[-1], "o", ms=5.5, color=c, mec=SURFACE, mew=1.4, zorder=4, clip_on=False)
    ax.set_xlim(x0, x1)
    lo, hi = min(ce.min(), cm.min()), max(ce.max(), cm.max())
    ax.set_ylim(np.floor((lo - 4) / 10) * 10, np.ceil((hi + 6) / 10) * 10)
    ax.yaxis.set_major_locator(MultipleLocator(10))
    ax.yaxis.set_major_formatter(FMT0)
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    style_axes(ax)
    y_top = ax.get_ylim()[1]
    for key in SUBPERIODS:
        p = ew["periods"][key]
        a, b = p["start"], p["end"] + pd.Timedelta(days=1)
        label = next(d for k, _, d in PERIODS if k == key)
        ax.text(a + (b - a) / 2, y_top - 1.8, label[0].upper() + label[1:], ha="center", va="top", fontsize=8.5, color=INK2, style="italic")
    ax.set_ylabel("Kumulált többlethozam (százalékpont)", fontsize=9)
    ax.annotate(f"Egyenlő súlyozás\n{hu(ce.iloc[-1])} pp", xy=(ce.index[-1], ce.iloc[-1]), xytext=(8, -6), textcoords="offset points",
                fontsize=8.5, va="center", color=INK, annotation_clip=False)
    ax.annotate(f"Momentum\n{hu(cm.iloc[-1])} pp", xy=(cm.index[-1], cm.iloc[-1]), xytext=(8, 4), textcoords="offset points",
                fontsize=8.5, va="center", color=INK, annotation_clip=False)
    fig.legend(handles=STRATEGY_LINES, loc="lower center", ncol=2, fontsize=8.5, bbox_to_anchor=(0.45, 0.0), handlelength=1.8)
    save(fig, out_dir, "abra_kumulalt_tobblethozam")


def _waterfall(ax, attr: pd.Series, color: str, title: str) -> None:
    vals = [float(attr[k]) for k in ORDER]
    running = np.cumsum(vals)
    prev = np.concatenate([[0.0], running[:-1]])
    w = 0.62
    halo = [pe.withStroke(linewidth=2.6, foreground=SURFACE)]
    for i, (a, b) in enumerate(zip(prev, running)):
        ax.bar(i, b - a, bottom=a, width=w, color=color, zorder=3)
        off = 2.2
        if vals[i] >= 0:
            ax.text(i, max(a, b) + off, hu(vals[i], 2, True), ha="center", va="bottom", fontsize=8, color=INK, zorder=7, path_effects=halo)
        else:
            ax.text(i, min(a, b) - off, hu(vals[i], 2), ha="center", va="top", fontsize=8, color=INK, zorder=7, path_effects=halo)
        if i < len(vals) - 1:
            ax.plot([i + w / 2, i + 1 - w / 2], [b, b], color=MUTED, lw=0.8, zorder=2)
    tot = running[-1]
    n = len(vals)
    ax.plot([n - 1 + w / 2, n - w / 2], [tot, tot], color=MUTED, lw=0.8, zorder=2)
    ax.bar(n, tot, width=w, color=INK2, zorder=3)
    ax.text(n, tot + (2.2 if tot >= 0 else -2.2), hu(tot, 2, True), ha="center", va="bottom" if tot >= 0 else "top",
            fontsize=8, fontweight="bold", color=INK, zorder=7, path_effects=halo)
    ax.axhline(0, color=MUTED, lw=0.9, zorder=4)
    ax.set_xlim(-0.6, n + 0.6)
    ax.set_xticks(range(n + 1))
    ax.set_xticklabels(XLABELS, fontsize=8.5)
    ax.tick_params(axis="x", length=0)
    ax.yaxis.set_major_locator(MultipleLocator(25))
    ax.yaxis.set_major_formatter(FMT0)
    style_axes(ax)
    ax.set_title(title, loc="left", fontsize=9, color=INK)
    ax.set_ylabel("Kumulált többlethozam (százalékpont)", fontsize=9)
    tr = ax.get_xaxis_transform()
    for a, b, t in ((0, 2, "egyedi"), (3, 5, "páros"), (6, 6, "hármas")):
        ax.plot([a - 0.3, b + 0.3], [-0.10, -0.10], color=BASELINE, lw=1.0, transform=tr, clip_on=False)
        ax.text((a + b) / 2, -0.135, t, transform=tr, ha="center", va="top", fontsize=8, color=INK2)


def fig_waterfall(ew: dict, mom: dict, out_dir: Path) -> None:
    """5. és 6. ábra: a hét osztalék vízesés-ábrája a teljes időszakra, külön az egyenlő súlyozásra és a momentumra (közös y-skálával)."""
    span = SHORT[FULL]
    runs = [np.cumsum([s["periods"][FULL]["attr"][k] for k in ORDER]) for s in (ew, mom)]
    lo, hi = min(r.min() for r in runs), max(max(r.max() for r in runs), 0.0)
    ylim = (lo - 18.5, np.ceil((hi + 20) / 25) * 25)
    for s, color, title, name in ((ew, C_EW, f"Egyenlő súlyozású aktív portfólió, {span}", "abra_osztalek_vizeses_egyenlo"),
                                  (mom, C_MOM, f"Momentum-stratégiájú aktív portfólió, {span}", "abra_osztalek_vizeses_momentum")):
        fig, ax = plt.subplots(figsize=(6.3, 3.4))
        fig.subplots_adjust(left=0.11, right=0.985, top=0.90, bottom=0.19)
        _waterfall(ax, s["periods"][FULL]["attr"], color, title)
        ax.set_ylim(*ylim)
        save(fig, out_dir, name)


def _shapley_panel(ax, ew: dict, mom: dict, key: str, big: bool) -> None:
    span = ax.get_ylim()[1] - ax.get_ylim()[0]
    off = 0.012 * span
    for gi, d in enumerate(DECISION_MAKERS):
        for k, (s, c) in enumerate(((ew, C_EW), (mom, C_MOM))):
            v = float(s["periods"][key]["shapley"][d])
            xc = gi + (-0.19 if k == 0 else 0.19)
            ax.bar(xc, v, width=0.36, color=c, zorder=3)
            ax.text(xc, v + (off if v >= 0 else -off), hu(v, 2, True), ha="center", va="bottom" if v >= 0 else "top",
                    fontsize=8 if big else 7.5, color=INK, rotation=0 if big else 90, zorder=5)
    ax.axhline(0, color=MUTED, lw=0.9, zorder=4)
    ax.set_xlim(-0.6, 2.6)
    ax.set_xticks(range(3))
    style_axes(ax)
    ax.tick_params(axis="x", length=0)


def fig_shapley(ew: dict, mom: dict, out_dir: Path) -> None:
    """6. ábra: a Shapley-értékek időszakonként, a két stratégiára."""
    fig = plt.figure(figsize=(6.3, 6.4))
    gs = fig.add_gridspec(2, 3, height_ratios=[1.25, 1], hspace=0.42, wspace=0.10, left=0.10, right=0.985, top=0.95, bottom=0.12)
    vals_full = [float(s["periods"][FULL]["shapley"][d]) for s in (ew, mom) for d in DECISION_MAKERS]
    vals_sub = [float(s["periods"][k]["shapley"][d]) for s in (ew, mom) for k in SUBPERIODS for d in DECISION_MAKERS]

    top = fig.add_subplot(gs[0, :])
    top.set_ylim(min(vals_full) - 18, max(vals_full) + 18)
    top.yaxis.set_major_locator(MultipleLocator(25))
    top.yaxis.set_major_formatter(FMT0)
    _shapley_panel(top, ew, mom, FULL, True)
    top.set_xticklabels(["Régió (R)", "Szektor (S)", "Szelekció (C)"], fontsize=9)
    top.set_title(f"{TITLE[FULL]}   |   v(N): egyenlő súlyozás {hu(ew['periods'][FULL]['vN'])}, momentum {hu(mom['periods'][FULL]['vN'])} pp",
                  loc="left", fontsize=9, color=INK)
    top.set_ylabel("Shapley-érték (százalékpont)", fontsize=9)
    top.legend(handles=STRATEGY_PATCHES, loc="upper left", fontsize=8.5, bbox_to_anchor=(0.0, 1.0), handlelength=1.2)

    axes = []
    for i, key in enumerate(SUBPERIODS):
        ax = fig.add_subplot(gs[1, i], sharey=axes[0] if axes else None)
        if not axes:
            ax.set_ylim(min(vals_sub) - 15.3, max(vals_sub) + 13.4)
            ax.yaxis.set_major_locator(MultipleLocator(10))
            ax.yaxis.set_major_formatter(FMT0)
            ax.set_ylabel("Shapley-érték (százalékpont)", fontsize=9)
        else:
            plt.setp(ax.get_yticklabels(), visible=False)
        _shapley_panel(ax, ew, mom, key, False)
        ax.set_xticklabels(["R", "S", "C"], fontsize=9)
        ax.set_title(f"{TITLE[key]}\nv(N): {hu(ew['periods'][key]['vN'], 2, True)} | {hu(mom['periods'][key]['vN'], 2, True)} pp",
                     loc="left", fontsize=8, color=INK)
        axes.append(ax)
    fig.text(0.10, 0.045, "v(N): a teljes többlethozam (a Shapley-értékek összege), sorrendben egyenlő súlyozás | momentum.",
             fontsize=8, color=INK2, ha="left", va="center")
    save(fig, out_dir, "abra_shapley_strategiak")


DIV_KEYS = ["REGION", "SECTOR", "SELECTION", "REGION + SECTOR", "REGION + SELECTION", "SECTOR + SELECTION", "REGION + SECTOR + SELECTION"]
DIV_LABELS = ["R", "S", "C", "RS", "RC", "SC", "RSC"]


def _dividend_panel(ax, ew: dict, mom: dict, key: str) -> None:
    allv = [float(s["periods"][key]["attr"][d]) for s in (ew, mom) for d in DIV_KEYS]
    span = (max(allv + [0]) - min(allv + [0])) or 1.0
    off = 0.015 * span
    for gi, d in enumerate(DIV_KEYS):
        for k, (s, c) in enumerate(((ew, C_EW), (mom, C_MOM))):
            v = float(s["periods"][key]["attr"][d])
            xc = gi + (-0.19 if k == 0 else 0.19)
            ax.bar(xc, v, width=0.36, color=c, zorder=3)
            ax.text(xc, v + (off if v >= 0 else -off), hu(v, 1, True), ha="center", va="bottom" if v >= 0 else "top",
                    fontsize=6.3, color=INK, rotation=90, zorder=5)
    ax.axhline(0, color=MUTED, lw=0.9, zorder=4)
    ax.set_xlim(-0.6, 6.6)
    ax.set_xticks(range(7))
    ax.set_xticklabels(DIV_LABELS, fontsize=8)
    lo, hi = min(allv + [0]), max(allv + [0])
    pad = 0.30 * (hi - lo)
    ax.set_ylim(lo - pad, hi + pad)
    ax.yaxis.set_major_formatter(FMT0)
    style_axes(ax)
    ax.tick_params(axis="x", length=0)


def fig_dividends(ew: dict, mom: dict, out_dir: Path) -> None:
    """7. abra: a het Harsanyi-osztalek a ket strategiara, a teljes idoszakra es a harom reszidoszakra (kulon skalan)."""
    fig, axes = plt.subplots(2, 2, figsize=(6.3, 5.6))
    fig.subplots_adjust(left=0.095, right=0.985, top=0.93, bottom=0.10, hspace=0.38, wspace=0.22)
    for ax, (key, short, desc) in zip(axes.ravel(), PERIODS):
        _dividend_panel(ax, ew, mom, key)
        ax.set_title(f"{short} ({desc})", loc="left", fontsize=9, color=INK)
    for ax in axes[:, 0]:
        ax.set_ylabel("Harsányi-osztalék (százalékpont)", fontsize=8.5)
    fig.legend(handles=STRATEGY_PATCHES, loc="lower center", ncol=2, fontsize=8.5, bbox_to_anchor=(0.54, 0.0), handlelength=1.2)
    save(fig, out_dir, "abra_osztalekok_strategiak")


def main() -> None:
    parser = argparse.ArgumentParser(description="A Harsányi-attribúció eredményeinek ábrái (3-7. ábra).")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help=f"kimeneti mappa (alapértelmezés: {DEFAULT_OUT})")
    args = parser.parse_args()
    out_dir: Path = args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    f_ew, f_mom = find_xlsx(PATTERN_EW), find_xlsx(PATTERN_MOM)
    print(f"Egyenlő súlyozás: {f_ew.relative_to(BASE)}\nMomentum:         {f_mom.relative_to(BASE)}")
    ew, mom = load_strategy(f_ew), load_strategy(f_mom)
    bench_na = benchmark_na_weights(ew)
    print("A hatékonysági azonosság minden időszakra és mindkét stratégiára teljesül.\nÁbrák készítése:")

    setup_style()
    fig_region_weights(ew, mom, bench_na, out_dir)
    fig_cumulative_excess(ew, mom, out_dir)
    fig_waterfall(ew, mom, out_dir)
    fig_shapley(ew, mom, out_dir)
    fig_dividends(ew, mom, out_dir)
    print("Kész.")


if __name__ == "__main__":
    main()
