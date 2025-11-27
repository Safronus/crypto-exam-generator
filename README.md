# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny a **libovolně hluboké podskupiny**, dva typy otázek (**klasická** / **BONUS**), **rich-text editor**, **multiselect**, **filtr**, import **DOCX** (Word).

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.7

- **Zarovnání**: přidána tlačítka **Vlevo / Na střed / Vpravo / Do bloku**. Funguje i na **vybranou část textu** (zarovná všechny odstavce v rozsahu výběru).
- **Vizuální odlišení ve stromu**: skupiny, podskupiny a otázky mají různé **standardní ikony** (systémové, macOS-friendly).
- Zachováno: náhled formátování, názvy otázek, DnD refresh, stromový přesun, import DOCX s odrážkami.

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

- **Aktuální verze:** `1.7` (release: 2025-11-27)
- Changelog (výběr):
  - `1.7` – zarovnání textu + ikony ve stromu.
  - `1.6a` – náhled formátování, fix přesunu a hromadného mazání.
  - `1.6` – názvy otázek (title) včetně importních defaultů.
  - `1.5b` – Qt6-safe formát, DnD refresh, stromový přesun.
  - `1.4` – zachování odrážek, ignorace A–F, oprava NameError.

---

## Git

```bash
git add main.py README.md
git commit -m "feat(editor): zarovnání textu; feat(ui): ikony ve stromu (v1.7)"
git tag v1.7
git push && git push --tags
```
