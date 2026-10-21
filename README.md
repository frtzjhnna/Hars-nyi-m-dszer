# Harsányi-féle teljesítményfelbontás részvényportfólióra

A repó a szakdolgozatomhoz készült Python-kódot tartalmazza. A kód egy 100 részvényes, globális
(észak-amerikai, európai, ázsiai) részvényportfólió benchmarkhoz viszonyított kumulált aktív hozamát bontja fel
**Harsányi-osztalékokra** három hierarchikus döntéshozó (régió **R**, szektor **S**, részvényszelekció **C**) szerint,
két aktív stratégiára:

* egyenlő súlyozású (hierarchikus: régiónként, régión belül szektoronként, cellán belül részvényenként egyenlő súly),
* momentum alapú (12–1 hónapos jel, rangalapú súlyeltolás mindhárom szinten, τ = 0,5, féléves újrasúlyozás).

A benchmark a piaci kapitalizációval súlyozott portfólió ugyanabból a 100 részvényből. A koalíciók értékét a
Davies–Laker-módszerrel (a koalíciós hipotetikus portfólió és a benchmark kumulált hozamának különbsége) számolja a kód,
a teljes időszakra (2016–2025) és három részidőszakra (2016–2019, 2020–2021, 2022–2025). A módszertani háttér:
Acerbi, Csóka és Herings (2026); Lundström és Hedbring (2025); Davies és Laker (2001).

## Szerkezet

```
kod/
  adatok_atalakitas.py                  nyers Bloomberg-export (adatok.xlsx) -> adatok_python.xlsx (Prices, Market_Cap, Metadata)
  véglegeshozamszámítás.ipynb           egyenlő súlyozású aktív portfólió: Harsányi-osztalékok, Shapley-értékek
  véglegeshozamszámítás_momentum.ipynb  momentum alapú aktív portfólió: Harsányi-osztalékok, Shapley-értékek
  axioma_ellenorzes.py                  hatékonyság, külső döntéshozótól való függetlenség, nulla hozzájárulás ellenőrzése
  hozzajarulas_bontas.py                az egyedi osztalékok bontása régió / szektor / részvény hozzájárulásra
  abrak_keszitese.py                    a dolgozat ábrái az eredményfájlokból
requirements.txt
```

## Adatok

A bemeneti adatok (havi árfolyamok és piaci kapitalizáció) **Bloomberg-lekérdezésből** származnak, ezért az adatfájlok
**nincsenek a repóban** (licenc). A kód futtatásához a repó gyökerébe kell tenni az `adatok.xlsx` fájlt:

* `Árfolyamok (2)` és `Piaci kap (2)` munkalap: tickerenként egy (dátum, érték) oszloppár, havi adatokkal
  (2015.12.31-től 2025.12.31-ig, a momentum-jelhez 2014.12.31-től),
* `Nevek` munkalap: a Bloomberg-tickerek és a vállalatnevek.

Az `adatok_atalakitas.py` ebből készíti a notebookok által várt `kod/adatok_python.xlsx` fájlt (`Prices`, `Market_Cap`,
`Metadata`, `README` munkalap). Ha a JNJ piaci kapitalizációja hiányzik az `adatok.xlsx`-ből, a szkript egy régi exportból
(`adatok_100reszveny_végleges.xlsx`) pótolja; a részleteket a szkript fejléce írja le.

## Futtatás

Python 3.11, a csomagok a `requirements.txt`-ben. A parancsokat a repó gyökeréből adjuk ki.

```bash
pip install -r requirements.txt

# 1. adatok átalakítása (adatok.xlsx a repó gyökerében)
python kod/adatok_atalakitas.py

# 2. a két stratégia felbontása: a notebookokat a kod mappából kell futtatni (relatív útvonalak),
#    mindkettőt végig kell futtatni (Run All)
cd kod
jupyter nbconvert --to notebook --execute --inplace "véglegeshozamszámítás.ipynb"
jupyter nbconvert --to notebook --execute --inplace "véglegeshozamszámítás_momentum.ipynb"

# 3. ellenőrzések, hozzájárulás-bontás és ábrák
python axioma_ellenorzes.py
python hozzajarulas_bontas.py
python abrak_keszitese.py            # kimenet: ../ábrák
```

A notebookok az `outputs_harsanyi_100stocks_*` mappákba írják az eredményfájlokat (Excel; osztalékok, Shapley-értékek, idősorok,
hatékonysági ellenőrzés). Az eredményfájlok a Bloomberg-adatokból számolt részvényszintű súlyokat és kapitalizációt is
tartalmaznak, ezért nem kerültek a repóba (`.gitignore`).

## Ellenőrzések

* A hét Harsányi-osztalék összege minden időszakban és mindkét stratégiánál pontosan kiadja a kumulált aktív hozamot.
* τ = 0 mellett a momentum alapú portfólió megegyezik az egyenlő súlyozásúval (`TILT_REGION`, `TILT_SECTOR`, `TILT_SELECTION` a momentum-notebookban).
* Az axióma-ellenőrzés (`axioma_ellenorzes.py`) a hatékonyságot, a külső döntéshozótól való függetlenséget és a nulla hozzájárulást teszteli.

## Hivatkozás

Szakdolgozat: <a Részvényportfólió elemzés Harsányi-módszerrel, 2026, Budapesti Corvinus Egyetem>.
