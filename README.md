# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny a **libovolně hluboké podskupiny**, dva typy otázek (**klasická** / **BONUS**), **rich-text editor**, **multiselect**, **filtr**, import **DOCX** (Word).

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.7b

- **BONUS body na 2 desetinná místa**: editor používá **QDoubleSpinBox** (krok 0.01), model ukládá **float**, ve stromu se zobrazuje `+X.XX/ Y.YY`.
- Zachováno: vizuální odlišení BONUS otázek, tučné skupiny, zarovnání textu, náhled formátování, DnD, import DOCX.

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

- **Aktuální verze:** `1.7b` (release: 2025-11-27)
- Changelog (výběr):
  - `1.7b` – BONUS body s přesností na dvě desetinná místa.
  - `1.7a` – vizuální odlišení BONUS vs. klasická + tučné skupiny.
  - `1.7` – zarovnání textu + ikony ve stromu.
  - `1.6a` – náhled formátování, fix přesunu a hromadného mazání.

---

## Git

```bash
git add main.py README.md
git commit -m "feat(bonus): dvě desetinná místa bodů; float v modelu; QDoubleSpinBox (v1.7b)"
git tag v1.7b
git push && git push --tags
```
