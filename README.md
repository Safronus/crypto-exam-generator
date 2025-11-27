# Crypto Exam Generator

Aplikace v **PySide6** pro správu zkušebních otázek (Kryptologie). Skupiny a **libovolně hluboké podskupiny**, dva typy otázek (**klasická** / **BONUS**), **rich-text editor**, **multiselect**, **filtr**, import **DOCX** (Word).

> Minimalisticky: **jeden soubor** `main.py` (GUI + logika).

---

## Novinky ve verzi 1.7d

- **Náhled formátování** skryt – editor je WYSIWYG, není potřeba duplikovat zobrazení.
- **Fix Tučného:** tučné písmo se aplikuje **hned při prvním kliku** na tlačítko/klávesovou zkratku (opravená logika akce).

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

- **Aktuální verze:** `1.7d` (release: 2025-11-27)
- Changelog (výběr):
  - `1.7d` – skrytý náhled + fix tučného.
  - `1.7c` – autosave všech změn otázky (debounce 1.2 s).
  - `1.7b` – BONUS body s přesností na dvě desetinná místa.

---

## Git

```bash
git add main.py README.md
git commit -m "fix(editor): bold na první klik; odstranit náhled (v1.7d)"
git tag v1.7d
git push && git push --tags
```
