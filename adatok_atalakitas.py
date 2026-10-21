"""
Az ../adatok.xlsx (nyers Bloomberg-export) átalakítása a notebookok által várt formátumra.

Kimenet: adatok_python.xlsx (ebben a mappában), munkalapok: Prices, Market_Cap, Metadata, README.
Az eredeti adatok.xlsx-et nem módosítja.

Használat (a "VS code" mappából):
    python adatok_atalakitas.py

Mit csinál:
  - az 'Árfolyamok (2)' és 'Piaci kap (2)' munkalapok (ticker-enként dátum+érték oszloppár) egyetlen
    Date-indexű táblává alakítása, hónap szerint illesztve (egyes ázsiai tickereknél a hónapvégi dátum 1-2 nappal eltér);
  - a Nevek munkalapból a Metadata (Bloomberg_Ticker, Company);
  - felfüggesztett kereskedés miatt hiányzó havi árak (pl. 600900 CH, 2015-07..10) kitöltése az utolsó jegyzett árral
    (a kitöltött cellák sárgák a Prices lapon, és a README-ben is szerepelnek);
  - ha a JNJ piaci kapitalizációja hiányzik az adatok.xlsx-ből, pótlás a régi exportból (lásd lent).
"""

import datetime as dt
import os

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "adatok.xlsx")
OLD_EXPORT = os.path.join(ROOT, "adatok_100reszveny_végleges.xlsx")
OUT = os.path.join(HERE, "adatok_python.xlsx")


def serial_to_datetime(x):
    return dt.datetime(1899, 12, 30) + dt.timedelta(days=int(x))


def block_frame(ws):
    """Ticker-enként (dátum, érték) oszloppár -> (év, hó) indexű DataFrame + a tickerek saját dátumai."""
    cols, dates = {}, {}
    for c in ws[1]:
        if c.value is None:
            continue
        values, own_dates = {}, {}
        for r in range(2, ws.max_row + 1):
            d = ws.cell(r, c.column).value
            v = ws.cell(r, c.column + 1).value
            if d is None:
                continue
            t = serial_to_datetime(d)
            key = (t.year, t.month)
            assert key not in values, f"duplikált hónap: {c.value} {key}"
            values[key] = v
            own_dates[key] = t
        cols[c.value] = pd.Series(values, dtype="float64")
        dates[c.value] = pd.Series(own_dates)
    return pd.DataFrame(cols), dates


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    sh_p = [n for n in wb.sheetnames if n.endswith("(2)") and n.startswith("Á")][0]
    sh_m = [n for n in wb.sheetnames if n.startswith("Piaci kap") and n.endswith("(2)")][0]

    P, Pd = block_frame(wb[sh_p])
    M, _ = block_frame(wb[sh_m])
    print(f"árfolyam: {P.shape}, {P.index[0]} -> {P.index[-1]} | piaci kap.: {M.shape}, {M.index[0]} -> {M.index[-1]}")
    assert not M.isna().any().any(), "hiányzó piaci kapitalizáció"
    assert P.index.is_monotonic_increasing

    filled_mask = P.isna()
    for t in P.columns[filled_mask.any()]:
        v = P[t]
        assert v.iloc[0] == v.iloc[0] and v.iloc[-1] == v.iloc[-1], f"{t}: az első/utolsó hónap hiányzik, nem tölthető ki"
    filled_cells = [(t, k) for t in P.columns for k in P.index[filled_mask[t]]]
    print("kitöltött cellák (felfüggesztett kereskedés):", filled_cells)
    P = P.ffill()
    assert not P.isna().any().any()
    month_no = P.index.map(lambda k: k[0] * 12 + k[1])
    assert (np.diff(month_no) == 1).all(), "az árfolyamsorban hiányzó hónap van"

    jnj_from_old = "JNJ US Equity" not in M.columns
    if jnj_from_old:
        old_p = pd.read_excel(OLD_EXPORT, sheet_name="Prices")
        old_m = pd.read_excel(OLD_EXPORT, sheet_name="Market_Cap")
        for df in (old_p, old_m):
            df.index = list(zip(df["Date"].dt.year, df["Date"].dt.month))
        shares = old_m["JNJ"].astype(float) / old_p["JNJ"].astype(float)
        shares_t = shares.shift(-1)
        shares_t.iloc[-1] = shares.iloc[-1]
        M["JNJ US Equity"] = P["JNJ US Equity"].reindex(M.index) * shares_t.reindex(M.index)
        assert M["JNJ US Equity"].notna().all()
    assert set(P.columns) == set(M.columns) and len(P.columns) == 100, "100 azonos tickert várok mindkét lapon"

    names_row = [c.value for c in wb["Nevek"][1] if c.value is not None]
    assert len(names_row) == 100
    rest = [c for c in P.columns if c != "JNJ US Equity"]
    name_of = dict(zip(rest[:5] + ["JNJ US Equity"] + rest[5:], names_row))

    order = ["JNJ US Equity"] + rest
    P, M = P[order], M[order]
    meta = pd.DataFrame({"Bloomberg_Ticker": order, "Company": [name_of[t] for t in order]})

    ref_dates = Pd["MSFT US Equity"]
    assert (ref_dates.index == P.index).all()
    for t in order:
        assert set(Pd[t].index) <= set(P.index), t
        assert set(P.index) - set(Pd[t].index) == {k for (tt, k) in filled_cells if tt == t}, t
    P_out = P.copy()
    P_out.index = pd.DatetimeIndex(ref_dates.values, name="Date")
    M_out = M.copy()
    M_out.index = pd.DatetimeIndex(ref_dates.reindex(M.index).values, name="Date")
    assert (P_out > 0).all().all() and (M_out > 0).all().all()
    shifted = sorted(t for t in order if any(Pd[t][k] != ref_dates[k] for k in Pd[t].index))

    readme = pd.DataFrame({
        "Információ": ["Fájl célja", "Forrás fájl", "Részvények száma", "Árfolyam időszak", "Piaci kapitalizáció időszak",
                       "Gyakoriság", "Pénznem", "Adatforrás", "Megjegyzés 1", "Megjegyzés 2", "Megjegyzés 3"],
        "Érték": [
            "Az adatok.xlsx Python-barát (notebookokkal futtatható) változata: Prices, Market_Cap, Metadata munkalapok. Előállítja: adatok_atalakitas.py",
            "adatok.xlsx: 'Árfolyamok (2)' és 'Piaci kap (2)' (beillesztett értékek) + 'Nevek' munkalap",
            "100",
            f"{P_out.index[0].date()} – {P_out.index[-1].date()}",
            f"{M_out.index[0].date()} – {M_out.index[-1].date()}",
            "havi", "USD", "Bloomberg (PX_LAST, CUR_MKT_CAP)",
            ("A JNJ piaci kapitalizációja az adatok.xlsx 'Piaci kap (2)' lapjáról hiányzik; pótlása: az adatok.xlsx JNJ-ára x a régi "
             "Bloomberg-export (adatok_100reszveny_végleges.xlsx) mkt cap / ár részvényszáma. A régi export JNJ-sora egy hónappal el volt "
             "csúszva (a t. hónap sora a t-1. hónapvégi értéket tartalmazza), ezért a részvényszámot a következő hónap sorából vettem; "
             "2025-12-hez az utolsó ismert részvényszám tartozik." if jnj_from_old
             else "Minden ticker minden adata az adatok.xlsx-ből származik."),
            "Eltérő hónapvégi dátumú tickerek (" + ", ".join(shifted) + "): az illesztés hónap alapján történt, az indexdátum az amerikai (MSFT) hónapvégi dátum.",
            ("Kitöltött árfolyamcellák (a Prices lapon sárgával jelölve): "
             + "; ".join(f"{t} {k[0]}-{k[1]:02d}" for t, k in filled_cells)
             + ". A Bloomberg ezekre a hónapokra nem ad árat (a kereskedés fel volt függesztve), ezért az utolsó jegyzett árat vittem tovább. "
               "A 12-1 momentum-jel csak a rangsor napja előtti 1. és 12. hónap árát használja (június/december rangsor-napoknál a "
               "november/május és az előző december/június), ezek nem kitöltöttek.") if filled_cells else "Nincs kitöltött cella.",
        ]})

    yellow = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")
    with pd.ExcelWriter(OUT, engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
        P_out.to_excel(xw, sheet_name="Prices")
        M_out.to_excel(xw, sheet_name="Market_Cap")
        meta.to_excel(xw, sheet_name="Metadata", index=False)
        readme.to_excel(xw, sheet_name="README", index=False)
        for t, k in filled_cells:
            xw.book["Prices"].cell(list(P.index).index(k) + 2, order.index(t) + 2).fill = yellow
        for ws in xw.book.worksheets:
            if ws.title in ("Prices", "Market_Cap"):
                ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 28 if ws.title == "README" else 18
            if ws.title == "README":
                ws.column_dimensions["B"].width = 150
    print("kész:", OUT)


if __name__ == "__main__":
    main()
