# MDLSZ Gyorspont

**MDLSZ Gyorspont** Windowsos asztali alkalmazás az MDLSZ gyorspont-lövészeti versenyek helyszíni adminisztrációjához. Segítségével versenyzőket kereshetsz és nevezhetsz, eredményeket rögzíthetsz, követheted a verseny állását, valamint kezelheted az élmezőny holtversenyeit.

**Aktuális kiadás: v1.0.2**

## Letöltés

[**MDLSZ_Gyorspont.exe letöltése Windowsra**](https://github.com/cscsanaki/MDLSZ_Gyorspont/releases/latest/download/MDLSZ_Gyorspont.exe)

[Összes kiadás és forráskód megtekintése](https://github.com/cscsanaki/MDLSZ_Gyorspont/releases)

## Főbb funkciók

- Versenyzők keresése versenyengedély-szám alapján a munkafüzet rejtett `Versenyengedelyek` lapjáról.
- Nevezés és nevezés visszavonása a tíz versenyszám bármelyikében.
- 0 és 195 közötti egész vagy tizedes eredmény, illetve fegyverhiba rögzítése.
- Az utoljára kezelt versenyzők, a teljes indulói névsor és a verseny statisztikáinak áttekintése.
- Versenyszámonkénti dobogó megjelenítése, holtverseny feloldása a 13, 12, 11, 10 és 9 találatok számossága alapján.
- A módosítások mentése a kiválasztott Excel-munkafüzetbe; a program megőrzi a cellastílusokat és képleteket.
- Biztonsági `.bak` másolat készítése a munkafüzet első megnyitásakor. A program bizonyos hibás XML-attribútumok javítását is megkísérli betöltéskor.

## Futtatás Pythonból

Windows alatt nyiss Parancssort a projekt mappájában, majd futtasd:

```bat
python -m pip install openpyxl
python mdlsz_gyorspont.py
```

## Önálló Windows `.exe` készítése

Telepített Python mellett kattints duplán a `build_exe.bat` fájlra, vagy futtasd Parancssorból. A build kézzel is elvégezhető:

```bat
python -m pip install openpyxl pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name "MDLSZ_Gyorspont" mdlsz_gyorspont.py
```

A kész fájl itt jön létre: `dist\MDLSZ_Gyorspont.exe`. Az `.exe` futtatásához a célgépen nem kell külön Pythont telepíteni.

## Használat versenymunkafüzettel

1. Indítsd el a programot, majd az **XLS betöltése** gombbal válaszd ki a hivatalos verseny `.xlsx` munkafüzetet.
2. Keresd meg a versenyzőt a versenyengedély-száma alapján.
3. Nevezd be a kívánt versenyszámba, majd a lőlap kiértékelése után rögzítsd az eredményt.
4. A **Statisztika** és az **Eredmények** fülön követheted a verseny állását, illetve rendezheted a holtversenyt.

A munkafüzetnek az alábbi szerkezetet kell követnie:

- `Versenyengedelyek` nevű lap: az A oszlopban a versenyengedély-szám, a B oszlopban a név, a C oszlopban az egyesület szerepel.
- Külön munkalap szükséges mind a tíz versenyszámhoz:
  - KKPI - Nyílt és KKPI - Optika
  - NKPI - Nyílt és NKPI - Optika
  - KKPU - Nyílt és KKPU - Optika
  - NKPU - Nyílt és NKPU - Optika
  - Sörétes - Nyílt és Sörétes - Optika
- A versenyzői adatok a 6. sortól kezdődnek. A program az alapértelmezett 50 sor fölé is bővíti az adatokat.

## Adatkezelés

A munkafüzet versenyzők nevét, engedélyszámát és egyesületét tartalmazhatja. Valódi versenyadatot, munkafüzetet, `.bak` másolatot vagy más személyes adatot ne tölts fel ebbe a repóba. Használat előtt készíts külön biztonsági másolatot, és először egy munkafüzet másolatával próbáld ki a programot. A program a kiválasztott munkafüzetet közvetlenül módosítja és menti.

## A repó fájljai

- `mdlsz_gyorspont.py` — a program Python-forráskódja
- `build_exe.bat` — Windowsos build parancsfájl
- `README.md` — projekt- és használati leírás

A lefordított `.exe` külön letölthető a fenti hivatkozásról. A kiadások oldalán a forráskód archívuma is elérhető.

## Licenc

Ehhez a repóhoz még nincs licenc kiválasztva. Licenc hozzáadásáig minden jog fenntartva; a szoftver felhasználásához, módosításához vagy továbbterjesztéséhez kérd a repó tulajdonosának engedélyét.
