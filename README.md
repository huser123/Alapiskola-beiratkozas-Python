# Alapiskola Beiratkozás Adatfeldolgozó

Ez a Python szkript .eml formátumú beiratkozási e-mailek adatainak feldolgozására szolgál. A program két fő célt szolgál:
1. Az összes beiratkozási e-mail adatainak exportálása CSV fájlba, hogy a beiratkozott tanulók adatai táblázatos formában is elérhetőek legyenek
2. Beiratkozási űrlapok generálása LaTeX formátumban, amely szép, nyomtatható dokumentumokat eredményez

## Telepítés

### Szükséges csomagok

A program a következő Python csomagokat használja:
- BeautifulSoup4 - HTML tartalom elemzéséhez
- codecs - karakterkódolás kezeléséhez

Telepítés:
```
pip install beautifulsoup4
```

A `codecs` a standard könyvtár része, nem kell külön telepíteni.

### Letöltés

```
git clone https://github.com/huser123/Alapiskola-beiratkozas-Python.git
cd Alapiskola-beiratkozas-Python
```

## Használat

1. Helyezd az összes feldolgozandó .eml fájlt ugyanabba a könyvtárba, ahol a szkript található
2. Készíts egy `adatlap.tex` fájlt, amely tartalmazza a LaTeX űrlap sablont a `!!!VARIABLE!!!` helyőrzőkkel
3. Futtasd a szkriptet:

```
python beiratkozas.py
```

A szkript két kimeneti fájlt hoz létre:
- `beiratkozasok.csv` - Az összes jelentkező adataival
- `beiratkozasok.tex` - Az összes jelentkező adatlapjával LaTeX formátumban

### A TeX sablon

A szkript egy TeX sablont használ, amelyben a `!!!VARIABLE!!!` helyőrzőket az aktuális adatokra cseréli. A helyőrzők sorrendje fontos, mivel ez határozza meg, hogy melyik mezőbe kerül az adott adat.

A sablonfájl (`adatlap.tex`) tartalmazzon teljes LaTeX dokumentumot a preambulummal és dokumentum lezárásával együtt. A program automatikusan csak egyszer használja a preambulumot és a lezáró részt.

## Mezők sorrendje

A rendszer a következő sorrendben helyettesíti be az adatokat:

1. A tanuló neve
2. A tanuló születési helye és dátuma
3. A tanuló születési száma
4. A tanuló állandó lakhelye
5. A tanuló állampolgársága
6. A tanuló nemzetisége
7. Az apa neve
8. Az apa e-mail címe
9. Az apa telefonszáma
10. Az apa állandó lakhelye
11. Az anya neve
12. Az anya e-mail címe
13. Az anya telefonszáma
14. Az anya állandó lakhelye
15. Melyik óvodába járt a gyermek
16. Milyen jellegű osztályt választana
17. Választható tantárgy
18. Napköziotthon
19. Iskolai étkeztetésre igényt tart
20. Van a gyermekének allergiája vagy más betegsége
21. A szülők egy háztartásban élnek
22. Elsődleges kapcsolattartási személy
23. Elsődleges kapcsolattartási telefonszám
24. Elsődleges kapcsolattartási e-mail cím
25. Az iskolalátogatásról szóló határozatot kinek címezheti
26. Bármilyen egyéb megjegyzés

## PDF generálás

A létrejött TeX fájlt egy LaTeX feldolgozóval (pl. pdflatex) lehet PDF-fé alakítani:

```
pdflatex beiratkozasok.tex
```

## Hibaelhárítás

### Karakterkódolási problémák

Ha karakterkódolási problémák lépnek fel, ellenőrizd a beérkező e-mailek kódolását. A szkript több módon is megpróbál különböző kódolásokat kezelni, de egyes különleges karakterek problémát okozhatnak.

### LaTeX hibák

Ha a TeX fájl generálása során hibák lépnek fel, ellenőrizd az adatlap.tex sablon fájl szintaktikai helyességét. A LaTeX érzékeny bizonyos speciális karakterekre (pl. %, $, _, {, }, stb.), amelyeket a szkript automatikusan próbál escape-elni.
