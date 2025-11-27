# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny a **libovolně hluboké podskupiny**, dva typy otázek (**klasická** / **BONUS**), **rich-text editor**, **multiselect**, **filtr**, import **DOCX** (Word).

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.5b

- **Opravy chyb s odsazením** (IndentationError) — celý soubor byl zrekonstruován tak, aby měl konzistentní odsazení.
- **Qt6 kompatibilita**: používá se `QFont.Weight.Bold/Normal`.
- **Drag & drop**: po přesunu se strom **znovu vykreslí** a přesunuté otázky se **znovu vyberou**, takže se **okamžitě** zobrazí jejich obsah.
- **Přesun vybraných / otázky**: dialog **stromu** pro výběr cíle (skupina/podskupina).
- **Import DOCX**: využití `word/numbering.xml`, ignorace škály **A→F**, zachování odrážek/číslování v HTML (`<ul>`, `<ol type="a">`, `<ol>`).

---

## Import z DOCX

- **Soubor → Import z DOCX…** (také tlačítko **Import** v toolbaru, zkratka **Ctrl/⌘+I**).
- Klasické otázky: každý **číslovaný odstavec** (úroveň 0) se importuje jako **samostatná otázka** s **1 bodem**.
- BONUS otázky: bloky `Otázka <číslo>` nebo text obsahující `BONUS`.
- **A→F stupnice** a administrativní texty se ignorují.
- Importované otázky se uloží do skupiny **„Neroztříděné“** (vytvoří se automaticky).

---

## Instalace (macOS)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install PySide6
python3 main.py
```

---

## Verze

- **Aktuální verze:** `1.6` (release: 2025-11-27)
- Changelog:
  - `1.5b` – rekonstrukce bez IndentationError, Qt6-safe formát, dnd refresh + reselection, stromový přesun.
  - `1.5` – dnd refresh, stromový přesun, Qt6 fix.
  - `1.4a` – tuple/dict kompatibilita při importu.
  - `1.4` – zachování odrážek, ignorace A–F, oprava NameError.
  - `1.3` – multiselect, filtr, opravy importu 1..10.

---

## Git

```bash
git add main.py README.md
git commit -m "fix!: rebuild to v1.5b (indentation, Qt6-safe, dnd refresh, tree move)"
git tag v1.5b
git push && git push --tags
```



---

## Novinky ve verzi 1.6

- **Název otázky**: každá otázka má editovatelný **název** (pole „Název otázky“ nad editorem). Název se zobrazuje i ve stromu.
- **Import DOCX**: nově se automaticky doplní **výchozí název** z prvního řádku/věty textu (pro BONUS s prefixem „BONUS:“). Název jde kdykoliv přepsat.
- **Kompatibilita**: starší JSON bez `title` se při načtení doplní (odvozením z textu).
