# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny a **libovolně hluboké podskupiny**, dva typy otázek (**klasická** / **BONUS**), **rich-text editor**, **multiselect**, **filtr**, import **DOCX** (Word).

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.7a

- **Vizuální odlišení typů otázek**: ve stromu jsou **BONUS** otázky zvýrazněny barvou a tučným písmem; klasické zůstávají standardní.
- **Skupiny** jsou nyní **tučně**, aby se jasně odlišily od podskupin.

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

- **Aktuální verze:** `1.7a` (release: 2025-11-27)
- Changelog (výběr):
  - `1.7a` – vizuální odlišení BONUS otázek + tučné skupiny.
  - `1.7` – zarovnání textu + ikony ve stromu.
  - `1.6a` – náhled formátování, fix přesunu a hromadného mazání.
  - `1.6` – názvy otázek (title) včetně importních defaultů.

---

## Git

```bash
git add main.py README.md
git commit -m "feat(ui): vizuální odlišení BONUS vs klasická + tučné skupiny (v1.7a)"
git tag v1.7a
git push && git push --tags
```
