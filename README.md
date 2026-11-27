# MDLSZ Gyorspont

**MDLSZ Gyorspont** Windowsos asztali alkalmazás az MDLSZ gyorspont lövészeti versenyek helyszíni adminisztrációjához. Segítségével versenyzőket kereshetsz és nevezhetsz, eredményeket rögzíthetsz, követheted a verseny állását, valamint kezelheted az élmezőny holtversenyeit.

> **Aktuális verzió: v1.0.6**

A kezdőképernyő az MDLSZ és a Gyorspont logóját jeleníti meg. A külön keretben található, félkövér **Frissítések keresése** gombbal ellenőrizhető, hogy jelent-e meg újabb kiadás a GitHubon. Új verzió esetén a program felajánlja a kiadáshoz tartozó Windowsos `.exe` megnyitását. Az ellenőrzés nem akadályozza az XLS betöltését, és internetkapcsolat nélkül is használható a program.

## Letöltés

[**MDLSZ_Gyorspont.exe letöltése Windowsra**](https://github.com/cscsanaki/MDLSZ_Gyorspont/releases/latest/download/MDLSZ_Gyorspont.exe)  
[Összes kiadás megtekintése](https://github.com/cscsanaki/MDLSZ_Gyorspont/releases)

## Főbb funkciók

- Versenyzők keresése versenyengedély-szám alapján a munkafüzet rejtett `Versenyengedelyek` lapjáról.
- Nevezés és nevezés visszavonása a tíz versenyszám bármelyikében.
- 0 és 195 közötti egész eredmény, illetve fegyverhiba rögzítése.
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
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "MDLSZ_Gyorspont" --add-data "assets;assets" mdlsz_gyorspont.py
```

A kész fájl itt jön létre: `dist\MDLSZ_Gyorspont.exe`. Ezt az `.exe`-t futtató Windows gépen nem kell külön Pythont telepíteni.

## Használat verseny munkafüzettel

1. Indítsd el a programot, majd az **XLS betöltése** gombbal válaszd ki a hivatalos verseny `.xlsx` munkafüzetet.
2. Keresd meg a versenyzőt a versenyengedély-száma alapján.
3. Nevezd be a kívánt versenyszámba, majd a lőlap kiértékelése után rögzítsd az eredményt.
4. A **Statisztika** és az **Eredmények** fülön követheted a verseny állását, illetve rendezheted a holtversenyt.

A munkafüzetnek az alábbi szerkezetet kell követnie:

- `Versenyengedelyek` nevű lap: A oszlopban versenyengedély-szám, B oszlopban név, C oszlopban egyesület.
- Külön munkalap mind a tíz versenyszámhoz:
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
- `assets/` — a kezdőképernyőn használt MDLSZ és Gyorspont logó
- `README.md` — projekt- és használati leírás

A lefordított `.exe` a GitHub **Releases** részében tölthető le.

## Licenc

Ehhez a repóhoz még nincs licenc kiválasztva. Licenc hozzáadásáig minden jog fenntartva; a szoftver felhasználásához, módosításához vagy továbbterjesztéséhez kérd a repó tulajdonosának engedélyét.

## Code signing policy

**Planned code signing:** Free code signing provided by SignPath.io, certificate by SignPath Foundation. The application is pending approval. Current release files are unsigned; signing will begin only after approval and workflow configuration.

**Project roles:** Csanaki Csaba (@cscsanaki) maintains the project, reviews proposed changes, and will approve release signing requests.

**Privacy:** The application processes the selected workbook locally and does not upload workbook or competitor data. The update check contacts GitHub only when the user starts the check.
